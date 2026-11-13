#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""배포 산출물 빌드 파이프라인 (게이트 통과 필수)

  문제: 세 번 연속 '변수가 정의되지 않았습니다' 가 났는데,
        매번 소스만 고치고 _src/횡단면도그리기.bas(CP949)와 zip 은
        다시 만들지 않아 사용자는 '고쳐지지 않은 파일'을 임포트했다.

  해결: 이 스크립트 하나만 실행하면 순서가 강제된다.
        1) gate.py 로 UTF-8 원본 검사        -> 실패하면 즉시 중단(exit 1)
        2) CP949 + CRLF 로 변환
        3) CP949 -> UTF-8 역변환 후 (1)과 '바이트 동일' 확인
        4) 역변환본에 gate.py 를 다시 실행     -> 인코딩으로 깨진 경우 차단
        5) mkzip.py 로 zip 생성 (게이트 재확인)
        6) 산출물 지문(SHA-256/MD5) 출력

  ※ 3)이 중요한 이유: CP949 에 없는 문자(예: —)는 조용히 '?' 로 바뀐다.
     그러면 UTF-8 원본은 통과하는데 실제 임포트본은 깨진 상태가 된다.

사용법:
    python3 _analysis/build.py
"""
import hashlib
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_UTF8 = os.path.join(ROOT, "_src", "횡단면도그리기_utf8.bas")
SRC_CP949 = os.path.join(ROOT, "_src", "횡단면도그리기.bas")
GATE = os.path.join(ROOT, "_analysis", "gate.py")
MKZIP = os.path.join(ROOT, "_analysis", "mkzip.py")


def hdr(t):
    print("\n" + "=" * 68)
    print(" " + t)
    print("=" * 68)


def run_gate(path, label):
    print("\n-- 게이트: %s --" % label)
    r = subprocess.run([sys.executable, GATE, path],
                       capture_output=True, text=True)
    print(r.stdout.rstrip())
    if r.returncode != 0:
        if r.stderr.strip():
            print(r.stderr.rstrip())
        print("\n*** 게이트 실패 — 빌드를 중단합니다. ***")
        return False
    return True


def main():
    hdr("1) UTF-8 원본 게이트")
    if not run_gate(SRC_UTF8, "횡단면도그리기_utf8.bas"):
        return 1

    hdr("2) CP949(ANSI) + CRLF 변환")
    text = open(SRC_UTF8, encoding="utf-8").read()
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # CP949 로 표현할 수 없는 문자가 있으면 미리 알려준다 (조용한 '?' 치환 방지)
    bad = sorted({c for c in text if not _can(c)})
    if bad:
        print("  ✗ CP949 로 표현 불가한 문자 %d종:" % len(bad))
        for c in bad:
            print("      %r  U+%04X" % (c, ord(c)))
        print("  => 이 문자를 제거/치환한 뒤 다시 실행하십시오.")
        return 1

    crlf = text.replace("\n", "\r\n")
    raw = crlf.encode("cp949")
    with open(SRC_CP949, "wb") as f:
        f.write(raw)
    print("  ✓ _src/횡단면도그리기.bas  %d B" % len(raw))

    hdr("3) 역변환 무손실 검증")
    back = open(SRC_CP949, "rb").read().decode("cp949")
    if back != crlf:
        print("  ✗ CP949 -> UTF-8 역변환이 원본과 다릅니다!")
        for i, (a, b) in enumerate(zip(crlf, back)):
            if a != b:
                print("      첫 차이 offset %d: %r != %r" % (i, a, b))
                break
        return 1
    print("  ✓ 왕복 완전 일치 (문자 단위)")

    hdr("4) CP949 역변환본 게이트 재검사")
    tmp = os.path.join(ROOT, "_verify", "_from_cp949_utf8.bas")
    os.makedirs(os.path.dirname(tmp), exist_ok=True)
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write(crlf)
    if not run_gate(tmp, "CP949 에서 되돌린 사본"):
        return 1

    hdr("5) 배포 zip 생성 (게이트 재확인 포함)")
    r = subprocess.run([sys.executable, MKZIP], capture_output=True, text=True)
    print(r.stdout.rstrip())
    if r.returncode != 0:
        if r.stderr.strip():
            print(r.stderr.rstrip())
        print("\n*** zip 생성 실패 ***")
        return 1

    hdr("6) 산출물 지문")
    targets = [
        SRC_UTF8,
        SRC_CP949,
        os.path.join(ROOT, "dist", "sangsu_cross_section_v2.4.zip"),
        os.path.join(ROOT, "dist", "sangsu_cross_section_v2.4_ko.zip"),
    ]
    for t in targets:
        if not os.path.exists(t):
            print("  ! 없음: %s" % os.path.basename(t))
            continue
        b = open(t, "rb").read()
        print("  %-34s %8d B  MD5 %s"
              % (os.path.basename(t), len(b), hashlib.md5(b).hexdigest()))

    print("\n완료. 위 MD5 를 업로드 후 다운로드 검증에 사용하십시오.")
    return 0


def _can(ch):
    try:
        ch.encode("cp949")
        return True
    except UnicodeEncodeError:
        return False


if __name__ == "__main__":
    sys.exit(main())
