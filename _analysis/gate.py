#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""빌드 게이트 — VBA '변수가 정의되지 않았습니다' 를 구조적으로 차단한다.

[이 게이트가 존재하는 이유]
세 번 연속 같은 오류가 났고, 내 검사기는 세 번 모두 "이상 없음" 이라고 거짓 보고했다.

  v2.1 : 선언 '위치' 를 아예 검사하지 않았다.
  v2.2 : Private Const AL_LEFT ...   -> (Const 키워드가 있어서) 검출됨
  v2.4 : Private gSlp As Double      -> (Dim 키워드가 없어서) ***검출 실패***

  기존 검사 정규식:  (?:Const|Dim|Static|Declare)  ... 키워드 '있을 때만' 선언으로 인정.
  그런데 VBA 선언문에는 키워드가 없는 형태가 얼마든지 있다.
      Private gSlp As Double        <- Dim 없음
      Private gLayName() As String  <- Dim 없음
  즉 '선언처럼 보이는 것' 을 정규식으로 추측하는 한 계속 놓친다.

[그래서 규칙을 뒤집었다]
  추측하지 않는다. VBA 문법상 '프로시저 밖' 에는 선언만 올 수 있는데,
  VBA 는 그 선언을 '첫 프로시저 이전' 에서만 유효하게 등록한다.
  따라서:

      첫 프로시저 이후 + 프로시저 본문 밖 + 비어있지 않은 줄  ==>  무조건 오류

  이 규칙은 줄의 생김새를 전혀 보지 않으므로, 어떤 형태의 선언이든 100% 잡는다.
  (주석 줄만 예외로 허용 — 주석은 컴파일 대상이 아니다.)

사용법:
    python3 _analysis/gate.py _src/횡단면도그리기_utf8.bas
    python3 _analysis/gate.py            # 기본: UTF-8 원본
종료코드: 0 = 통과, 1 = 실패 (mkzip.py 가 이 코드를 확인하고 빌드를 중단한다)
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT = os.path.join(ROOT, "_src", "횡단면도그리기_utf8.bas")

# 프로시저 시작 / 끝.  한글 프로시저명도 있으므로 [^\s(] 로 넓게 잡는다.
RE_PROC = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(Sub|Function)\s+([^\s(]+)", re.I)
RE_PROP = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(Property\s+(?:Get|Let|Set))\s+([^\s(]+)", re.I)
RE_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\s*$", re.I)


def strip_comment(s):
    """문자열 리터럴 밖의 ' 주석을 제거."""
    out, q = [], False
    for ch in s:
        if ch == '"':
            q = not q
            out.append(ch)
        elif ch == "'" and not q:
            break
        else:
            out.append(ch)
    return "".join(out)


def logical_lines(raw_lines):
    """_ 로 이어지는 줄을 하나의 논리 줄로 병합. (줄번호는 시작 줄 유지)"""
    out, buf, start = [], "", None
    for i, ln in enumerate(raw_lines, 1):
        t = ln.rstrip()
        if buf == "":
            start = i
        # 주석만 있는 줄은 병합 대상이 아니다
        body = strip_comment(t).rstrip()
        if body.endswith("_") and not body.lstrip().startswith("'"):
            buf += body[:-1] + " "
            continue
        buf += t
        out.append((start, buf))
        buf = ""
    if buf:
        out.append((start, buf))
    return out


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT
    if not os.path.exists(path):
        print("파일 없음:", path)
        return 1

    raw = open(path, encoding="utf-8", errors="replace").read()
    src = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    logical = logical_lines(src)

    # ── 프로시저 구간 파악 ──────────────────────────────────────────────
    procs, cur, depth = [], None, 0
    for ln, txt in logical:
        s = strip_comment(txt).strip()
        if cur is None:
            m = RE_PROC.match(s) or RE_PROP.match(s)
            if m:
                cur = {"name": m.group(2), "start": ln, "end": None}
                procs.append(cur)
        else:
            # 중첩 Sub/Function 은 VBA 에 없다. End xxx 만 본다.
            if RE_END.match(s):
                cur["end"] = ln
                cur = None
    for p in procs:
        if p["end"] is None:
            p["end"] = 10 ** 9

    if not procs:
        print("*** 프로시저를 찾지 못했습니다. 파일 형식을 확인하십시오.")
        return 1

    first = procs[0]
    inside = set()
    for p in procs:
        for ln in range(p["start"], min(p["end"], len(src)) + 1):
            inside.add(ln)

    # ── 핵심 검사 ───────────────────────────────────────────────────────
    bad = []
    for ln in range(first["start"] + 1, len(src) + 1):
        if ln in inside:
            continue
        s = strip_comment(src[ln - 1]).strip()
        if not s:
            continue
        bad.append((ln, src[ln - 1].rstrip()))

    print("=" * 68)
    print(" VBA 선언부 게이트 : %s" % os.path.basename(path))
    print("=" * 68)
    print(" 프로시저 %d개 / 첫 프로시저 L%d (%s) / 끝 L%s"
          % (len(procs), first["start"], first["name"],
             procs[-1]["end"] if procs[-1]["end"] != 10 ** 9 else "-"))
    print()
    print(" 규칙 : 첫 프로시저 **이후** + 프로시저 **밖** + 비어있지 않은 줄 = 오류")
    print("        (줄의 생김새를 보지 않으므로 Dim 없는 선언도 잡힌다)")
    print()

    if bad:
        print(" ✗ 위반 %d건 — 이대로 임포트하면 100%% 컴파일 오류가 납니다:" % len(bad))
        for ln, t in bad[:40]:
            print("     L%-6d | %s" % (ln, t[:80]))
        if len(bad) > 40:
            print("     ... 외 %d건" % (len(bad) - 40))
        print()
        print(" => 위 줄들을 첫 프로시저(L%d) **위**로 옮기십시오." % first["start"])
        return 1

    print(" ✓ 위반 0건 — 프로시저 밖에 선언이 남아있지 않습니다.")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
