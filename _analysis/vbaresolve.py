#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VBA 컴파일 오류 검사기 (vbaresolve) — 최종판

[이 검사기가 존재하는 이유]
"컴파일 오류입니다: 변수가 정의되지 않았습니다" 는 두 가지로 발생한다.

  (A) 선언이 아예 없다.
  (B) 선언은 있는데 '프로시저 뒤'에 있어 VBA 가 보지 못한다.   <-- 진짜 함정

(B) 는 '미선언' 이 아니라 '위치' 문제라서, 선언 유무만 보는 검사기는 못 잡는다.
이 프로젝트에서 실제로 두 번 이 유형으로 실패했다.

  v2.1 : Private Const AL_LEFT ... 5개가 xcPolyA 뒤   (Const 키워드 O -> 검출됨)
  v2.4 : Private gSlp As Double 가 Wz 뒤              (Dim 키워드 X -> 검출 실패)

  *** 'Private gSlp As Double' 에는 Dim 이 없다 ***
  이전 검사 정규식이 (Const|Dim|Static|Declare) 를 요구했기 때문에
  이 줄을 선언으로 인식조차 하지 못했다. 그래서 계속 "이상 없음" 이 나왔다.

[이 검사기가 이전과 다른 점]
  1. 라인 컨티뉴에이션(_)을 먼저 병합한 '논리 줄' 로 판정한다.
  2. 프로시저 본문 안/밖을 추적한다 (들여쓰기 추측 아님).
  3. 본문 밖 선언은 Dim 없이 'Private x As T' 형태도 모두 잡는다.
  4. 파라미터는 시그니처 '전체'(병합된 논리 줄)에서 콤마 분해해 수집한다.
  5. 선언 이름은 콤마로 분해해 전부 수집한다 (Private a As T, b As T).
  6. 프로시저 이름도 선언된 이름으로 취급한다.
  7. 내장함수/상수/키워드/문장시작어 목록을 넉넉히 둔다.

사용법:
    python3 _analysis/vbaresolve.py
    python3 _analysis/vbaresolve.py path/to/file.bas
"""
import os
import re
import sys

DEFAULT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "_src", "횡단면도그리기_utf8.bas")


def strip_code(line):
    """주석 제거 (문자열 리터럴 안의 ' 는 보존)."""
    out, inq = [], False
    for ch in line:
        if ch == '"':
            inq = not inq
            out.append(ch)
        elif ch == "'" and not inq:
            break
        else:
            out.append(ch)
    return "".join(out)


def logical_lines(raw_lines):
    """라인 컨티뉴에이션(_) 병합 -> [(시작줄번호, 병합텍스트)]"""
    out, i = [], 0
    while i < len(raw_lines):
        start = i + 1
        buf = raw_lines[i]
        while re.search(r"\s_\s*$", strip_code(buf)) and i + 1 < len(raw_lines):
            buf = re.sub(r"\s_\s*$", " ", buf)
            i += 1
            buf += raw_lines[i].strip()
        out.append((start, buf))
        i += 1
    return out


RE_PROC = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?"
    r"(Sub|Function)\s+([^\s(]+)", re.I)
RE_PROP = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?Property\s+(Get|Let|Set)\s+([^\s(]+)", re.I)
RE_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\s*$", re.I)

# 모듈레벨/지역 선언 판정 — Dim 키워드가 없어도 잡는다.
RE_DECL_MOD = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Global)\s+|Dim\s+|Static\s+)?"
    r"(?:Const\s+|Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+)?"
    r"[A-Za-z_\uac00-\ud7a3]\w*\s*(?:\(\))?\s*As\s+\S", re.I)
RE_DECL_KEYWORD = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Const|Dim|Static|Declare)\b", re.I)

KEYWORDS = set("""
and as attribute byref byval call case close const declare dim do each else elseif end
endif endsub endfunction endwith endselect endtype endenum enum eqv erase error event exit
explicit false for friend function get global gosub goto if implements imp implements in is
let like loop lset me mod new next not nothing null on option optional or paramarray
preserve private property public raiseevent redim rem resume return rset select set static
step stop sub then to true type typeof until wend while with withevents write xor
""".split())

BUILTIN = set("""
abs array asc ascw atn callbyname cbool cbyte ccur cdate cdbl cdec chr chrw cint clng cos
createobject csng cstr cvar cverr curdir cvdbl cverr date dateadd datediff datepart
dateserial datevalue day ddb dir doevents environ eof error exp fileattr filedatetime
filelen filter fix format freefile fv getall getattr getobject hex hour iif imestatus
input instr instrrev int ipmt irr isarray isdate isempty iserror ismissing isnull isnumeric
isobject join lbound lcase left len loc lof log ltrim mid minute mirr month monthname msgbox
now npv oct partition pmt ppmt pv qbcolor raise rate replace rgb right rnd round rtrim
second seek sgn shell sin slk space spc split sqr str strcomp strconv string strreverse
switch sydb syscmd tab tan time timer timeserial timevalue trim typename ubound ucase val
vartype weekday weekdayname year appactivate beep chdrive chdir chdrive filecopy inputbox
kill load mkdir put randomize reset rmdir savepicture sendkeys setattr unload savestting
getsetting savesetting deleteetting debug print collection dictionary
""".split())

SYSTEM_OBJECTS = set("""
application thisworkbook activeworkbook activewindow activesheet activecell selection
worksheets workbooks range cells rows columns names debug err excel
""".split())

VB_CONST = set("""
vbcr vbcrlf vbcrlf vblf vbtab vbnullchar vbnullstring vbback vbformfeed vbverticaltab
vbobjecterror vbokonly vbokcancel vbyesno vbyesnocancel vbyes vbno vbcancel vbabort
vbretry vbignore vbcritical vbinformation vbquestion vbexclamation vbdefaultbutton1
vbdefaultbutton2 vbdefaultbutton3 vbdefaultbutton4 vbapplicationmodal vbsystemmodal
vbhide vbnormal vbnormalfocus vbmaximizedfocus vbminimizedfocus vbminimizednofocus
vbbinarycompare vbtextcompare vbpressandrelease vbldropdownlist vbldropdowncombo
vbleftbutton vbrightbutton vbmiddlebutton vboletrue vbolefalse vbget vblet vbset
""".split())

XL_CONST = set("""
xlup xldown xltoleft xltoright xlnone xlcontinuous xlthin xlmedium xlthick xldouble
xlcenter xlleft xlright xltop xlbottom xlautomatic xloutside xlinside xlgeneral xlvalue
xlwhole xlpart xlbyrows xlbycolumns xlascending xldescending xlyes xlno xlcancel
xlerrordialog xlcalculationautomatic xlcalculationmanual xlopenxmlworkbook xlsolid
xlpattern xldown xltoright xlformcontrolsnone xlvalidalertstop xlvalidalertwarning
xlbetween xlnotbetween xlequal xlgreater xlless
""".split())

RE_IDENT = re.compile(r"(?<![.\w\uac00-\ud7a3])([A-Za-z_\uac00-\ud7a3][\w\uac00-\ud7a3]*)")


def decl_names(text):
    """선언 줄에서 선언되는 이름들을 콤마 분해로 모두 뽑는다."""
    s = strip_code(text)
    s = re.sub(r"^\s*(?:(?:Public|Private|Friend|Global|Dim|Static)\s+)+", "", s, flags=re.I)
    s = re.sub(r"^Const\s+", "", s, flags=re.I)
    if re.match(r"^Declare\b", s, re.I):
        m = re.search(r"(?:Sub|Function)\s+(\w+)", s, re.I)
        return [m.group(1)] if m else []
    if re.match(r"^(Type|Enum)\s+\w", s, re.I):
        m = re.search(r"(?:Type|Enum)\s+(\w+)", s, re.I)
        return [m.group(1)] if m else []
    names = []
    for part in s.split(","):
        m = re.match(r"\s*([A-Za-z_\uac00-\ud7a3]\w*)", part)
        if m:
            names.append(m.group(1))
    return names


def param_names(sig):
    """프로시저 시그니처(병합된 논리 줄)에서 파라미터 이름을 뽑는다."""
    if "(" not in sig:
        return []
    ps = sig[sig.index("("):]
    if ")" in ps:
        ps = ps[:ps.rfind(")")]
    out = []
    for part in ps.split(","):
        # [수정] 'Optional ByVal x' 처럼 Optional 이 ByVal 앞에 온다.
        #        예전 정규식은 ByVal 을 Optional 앞에 요구해서, 그런 파라미터는
        #        이름 대신 'ByVal' 을 잡고 제외목록에 걸려 통째로 사라졌다.
        #        -> xcText 의 al/rot/rotGiven, xcPoly 의 bClosed 가 미선언으로 오탐.
        m = re.search(
            r"(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)*"
            r"([A-Za-z_\uac00-\ud7a3]\w*)", part)
        if m and m.group(1).lower() not in ("byval", "byref", "optional", "paramarray"):
            out.append(m.group(1))
    return out


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT
    if not os.path.exists(path):
        print("파일 없음:", path)
        return 2

    raw = open(path, encoding="utf-8", errors="replace").read()
    src = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    logical = logical_lines(src)

    # ── 프로시저 구간 ────────────────────────────────────────────────────
    procs, cur = [], None
    for ln, txt in logical:
        s = strip_code(txt)
        m = RE_PROC.match(s) or RE_PROP.match(s)
        if m:
            kind = (m.group(1) or "property").lower()
            cur = {"name": m.group(2), "start": ln, "end": None, "kind": kind,
                   "sig": s}
            procs.append(cur)
            continue
        if RE_END.match(s) and cur is not None:
            cur["end"] = ln
            cur = None
    for p in procs:
        if p["end"] is None:
            p["end"] = 10 ** 9

    first_proc = procs[0]["start"] if procs else 10 ** 9
    bounds = [(p["start"], p["end"]) for p in procs]

    def owner(ln):
        for p in procs:
            if p["start"] <= ln <= p["end"]:
                return p
        return None

    print("파일      : %s" % os.path.basename(path))
    print("논리 줄   : %d  (원본 %d 줄)" % (len(logical), len(src)))
    print("프로시저  : %d 개   첫 프로시저 L%d %s"
          % (len(procs), first_proc, procs[0]["name"] if procs else "-"))

    # ── 1) 모듈레벨 선언 수집 + 위치 검사 ────────────────────────────────
    mod_decl, misplaced, dup = [], [], []
    seen = {}
    for ln, txt in logical:
        if owner(ln) is not None:
            continue
        s = strip_code(txt).strip()
        if not s:
            continue
        if not (RE_DECL_KEYWORD.match(txt) or RE_DECL_MOD.match(txt)):
            continue
        names = decl_names(txt)
        for n in names:
            mod_decl.append((ln, txt.strip(), n))
            k = n.lower()
            if k in seen:
                dup.append((n, seen[k], ln))
            else:
                seen[k] = ln
        if ln > first_proc:
            misplaced.append((ln, txt.strip(), names))

    print("\n=== 1. 모듈레벨 선언 위치 (가장 중요) ===")
    print("   모듈레벨 선언 %d개, 고유 이름 %d개" % (len(mod_decl), len(seen)))
    if misplaced:
        print("   ✗ 첫 프로시저(L%d) **뒤**에 선언 %d건 — VBA 가 보지 못한다!"
              % (first_proc, len(misplaced)))
        for ln, t, ns in misplaced:
            print("      L%-6d %s" % (ln, t[:74]))
    else:
        print("   ✓ 모두 선언부(첫 프로시저 이전) — 정상")

    print("\n=== 2. 모듈레벨 중복 선언 ===")
    if dup:
        for n, a, b in dup:
            print("   ✗ '%s'  L%d / L%d" % (n, a, b))
    else:
        print("   ✓ 없음")

    # ── 2) 선언된 이름 집합 ──────────────────────────────────────────────
    declared = set(seen)
    for p in procs:
        declared.add(p["name"].lower())
        for pn in param_names(p["sig"]):
            declared.add(pn.lower())
    # 지역 선언
    for ln, txt in logical:
        if owner(ln) is None:
            continue
        s = strip_code(txt).strip()
        if re.match(r"^(Dim|Static|Const)\b", s, re.I):
            for n in decl_names(txt):
                declared.add(n.lower())

    # ── 3) 본문 식별자 대조 ──────────────────────────────────────────────
    unresolved = {}
    for ln, txt in logical:
        p = owner(ln)
        if p is None:
            continue
        s = strip_code(txt).strip()
        if not s:
            continue
        if re.match(r"^(Dim|Static|Const|ReDim|Private|Public)\b", s, re.I):
            continue
        if ln == p["start"] and "(" in s:      # 시그니처 줄은 파라미터만
            continue
        if re.match(r"^[A-Za-z_]\w*:\s*$", s):  # 라벨
            continue
        s = re.sub(r'"[^"]*"', '""', s)
        s = re.sub(r"\bAs\s+[\w\.]+", " ", s, flags=re.I)
        s = re.sub(r"^On\s+Error\s+GoTo\s+\w+", " ", s, flags=re.I)
        s = re.sub(r"\bGoTo\s+\w+", " ", s, flags=re.I)
        s = re.sub(r"\bGoSub\s+\w+", " ", s, flags=re.I)
        for m in RE_IDENT.finditer(s):
            nm = m.group(1)
            k = nm.lower()
            if k in declared or k in KEYWORDS or k in BUILTIN:
                continue
            if k in VB_CONST or k in XL_CONST or k in SYSTEM_OBJECTS:
                continue
            if k.startswith("vb") or k.startswith("xl"):
                continue
            if s[m.end():].lstrip().startswith(":="):   # 명명된 인자
                continue
            unresolved.setdefault(k, []).append((ln, p["name"], nm))

    print("\n=== 3. 미선언 식별자 (Option Explicit) ===")
    if unresolved:
        print("   ✗ %d종" % len(unresolved))
        for k in sorted(unresolved):
            ln, pn, orig = unresolved[k][0]
            print("      %-16s %2d회  첫 사용 L%-6d in %s"
                  % (orig, len(unresolved[k]), ln, pn))
    else:
        print("   ✓ 없음")

    # ── 4) 변수/프로시저 이름 충돌 ───────────────────────────────────────
    print("\n=== 4. 변수명 ↔ 프로시저명 충돌 ===")
    procset = set(p["name"].lower() for p in procs)
    clash = [(n, ln) for ln, t, n in mod_decl if n.lower() in procset]
    if clash:
        for n, ln in clash:
            print("   ✗ '%s' (변수 L%d) 가 프로시저 이름과 같음" % (n, ln))
    else:
        print("   ✓ 없음")

    bad = len(misplaced) + len(dup) + len(unresolved) + len(clash)
    print("\n" + "=" * 64)
    if bad == 0:
        print("결과: 문제 없음")
        return 0
    print("결과: 위치 %d / 중복 %d / 미선언 %d / 충돌 %d  = %d건"
          % (len(misplaced), len(dup), len(unresolved), len(clash), bad))
    return 1


if __name__ == "__main__":
    sys.exit(main())
