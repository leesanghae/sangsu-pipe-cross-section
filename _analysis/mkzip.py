#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""배포용 zip 생성기 — Windows 탐색기에서 '빈 폴더' 로 보이는 문제 방지

[왜 이 스크립트가 필요한가]
Linux 의 `zip` (Info-ZIP) 은 한글 파일명을 UTF-8 바이트로 저장하면서도
일반플래그(general purpose bit flag) 0x800 을 세우지 않는다.
Windows 탐색기는 0x800 이 없으면 파일명을 '현재 ANSI 코드페이지'(한국어 = CP949)
로 해석하는데, UTF-8 바이트(예: ed 9a a1 ...)는 CP949 로 디코딩되지 않으므로
해당 항목을 통째로 건너뛴다 -> "이 폴더는 비어 있습니다".

[해결]
1순위 : 파일명을 전부 ASCII 로 만든다  -> 어떤 OS/압축도구에서도 100% 안전
2순위 : 한글 이름을 유지해야 하면 Python zipfile 로 만들어
        자동으로 붙는 0x800 플래그에 의존한다.

VBA 임포트는 파일 '이름' 이 아니라 파일 내부의
`Attribute VB_Name = "횡단면도그리기"` 를 사용하므로
ASCII 이름(cross_section_draw.bas)으로 배포해도 모듈 이름은 동일하다.

사용법:
    python3 _analysis/mkzip.py
"""
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# (원본경로, ASCII 이름, 한글 이름)
FILES = [
    ("_src/횡단면도그리기.bas",       "cross_section_draw.bas",      "횡단면도그리기.bas"),
    ("_src/횡단면도그리기_utf8.bas",  "cross_section_draw_utf8.bas", "횡단면도그리기_utf8.bas"),
    ("횡단면도_사용절차서.md",         "MANUAL_ko.md",                "횡단면도_사용절차서.md"),
    ("횡단면도_분석검증_보고서.md",     "ANALYSIS_REPORT_ko.md",       "횡단면도_분석검증_보고서.md"),
]

README = """\
상수관로 횡단면도 자동작도 - 배포 패키지 (v2.4)
================================================

[v2.4 에서 고친 것]
  '컴파일 오류입니다: 변수가 정의되지 않았습니다' (gSlp 강조 표시)

  원인: 아래 두 선언이 첫 프로시저(InitParams)보다 **뒤**에 있었습니다.
            Private gCurK As Long          <- 원래 1310행
            Private gSlp As Double         <- 원래 1311행
        VBA 는 모듈 레벨 선언을 '첫 프로시저 이전' 에서만 인식합니다.
        그래서 Wz() 안의 gSlp 를 미정의 변수로 판단했습니다.

  수정: 두 줄을 선언부(71~72행)로 옮겼습니다. 첫 프로시저는 127행입니다.

  확인 방법(임포트 후 30초):
        Alt+F11 -> Ctrl+F -> "gSlp As Double" 검색
        결과가 1건이고 그 줄번호가 72 근처면 정상입니다.
        (1300번대에서 나오면 이전 파일을 임포트한 것입니다.)


1. VBA 에 임포트할 파일
------------------------
    cross_section_draw.bas        <-- 이것을 임포트 (CP949/ANSI, 실제 사용본)
    cross_section_draw_utf8.bas   <-- 같은 내용의 UTF-8 원본 (열어서 확인용)

  * 임포트는 반드시 cross_section_draw.bas 로 하십시오.
    (utf8 판을 임포트하면 한글이 깨집니다.)
  * 파일 이름은 임포트 결과에 영향이 없습니다.
    VBA 모듈 이름은 파일 내부의 Attribute VB_Name = "횡단면도그리기" 로 정해집니다.


2. 문서
-------
    MANUAL_ko.md            사용 절차서
    ANALYSIS_REPORT_ko.md   분석 검증 보고서


3. 임포트 순서
--------------
    1) Excel 에서 Alt + F11
    2) 기존 '횡단면도그리기' 모듈이 있으면 우클릭 -> 제거 (저장: 아니오)
    3) 파일 > 파일 가져오기 -> cross_section_draw.bas
    4) 디버그 > VBAProject 컴파일

  * 기존 모듈을 지우지 않고 다시 임포트하면 '횡단면도그리기1' 이 생겨
    이름이 중복되며 컴파일 오류가 재발합니다. 반드시 먼저 제거하십시오.
  * Module1 / Module2 (종단면도) 는 그대로 두셔도 됩니다.


4. 실행 순서
------------
    (최초 1회)  Alt+F8 -> 횡단_P시트생성
                Alt+F8 -> 횡단_1단계_검증      (캐드 없이 데이터만 확인)
                Alt+F8 -> 횡단_작도            (기준점 클릭 -> 일률 작도)


5. 문제가 생기면
----------------
  * '변수가 정의되지 않았습니다'  -> 기존 모듈을 제거하지 않고 다시 임포트한 경우입니다.
                                     모듈을 지우고 다시 임포트하십시오.
  * '시트를 찾을 수 없습니다'      -> P 시트의 SH_TOKJUK / SH_SDBH 값을 실제 시트명으로.
  * 캐드에 안 그려짐              -> 횡단_도곽확인 으로 블록/도곽 상태를 먼저 점검.
"""


def build(entries, out_path, label):
    """entries: [(src_abs, arcname), ...]"""
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("README.txt", README)
        for src, arc in entries:
            if not os.path.exists(src):
                print("  ! 없음:", src)
                continue
            z.write(src, arcname=arc)

    print("\n[%s]  %s" % (os.path.basename(out_path), label))
    ok = True
    with zipfile.ZipFile(out_path) as z:
        for i in z.infolist():
            is_ascii = all(ord(c) < 128 for c in i.filename)
            has_flag = bool(i.flag_bits & 0x800)
            if is_ascii:
                state = "ASCII (모든 OS 안전)"
            elif has_flag:
                state = "UTF-8 플래그 OK"
            else:
                state = "*** 플래그 없음 -> Windows 에서 깨짐 ***"
                ok = False
            print("   %-34s %8d B  %s" % (i.filename, i.file_size, state))
    print("   =>", "정상" if ok else "문제 있음")
    return ok


def main():
    dist = os.path.join(ROOT, "dist")
    all_ok = True

    # 1순위 : 전부 ASCII 이름 — 어떤 환경에서도 안전
    all_ok &= build(
        [(os.path.join(ROOT, s), a) for s, a, _k in FILES],
        os.path.join(dist, "sangsu_cross_section_v2.4.zip"),
        "1순위 권장 (파일명 영문)")

    # 2순위 : 한글 이름 유지 (UTF-8 플래그)
    all_ok &= build(
        [(os.path.join(ROOT, s), k) for s, _a, k in FILES],
        os.path.join(dist, "sangsu_cross_section_v2.4_ko.zip"),
        "2순위 (한글 파일명, UTF-8 플래그)")

    if not all_ok:
        sys.exit(1)
    print("\n완료.")


if __name__ == "__main__":
    main()
