#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VBA 정적 구조 검사기 v2 (라인컨티뉴에이션 병합 후 블록 균형 / 중복 Case / Array 크기)"""
import re, collections

PATH = "_src/횡단면도그리기_utf8.bas"
raw = open(PATH, encoding="utf-8").read().split("\n")


def strip_code(s):
    out, inq = [], False
    for c in s:
        if c == '"':
            inq = not inq
            out.append(c)
        elif c == "'" and not inq:
            break
        else:
            out.append(c)
    return "".join(out)


# ---- 라인 컨티뉴에이션 병합 (원본 줄번호 유지) ----
logical = []          # (lineno_of_start, merged_text)
i = 0
while i < len(raw):
    start = i + 1
    buf = raw[i]
    while re.search(r"\s_\s*$", strip_code(buf)) and i + 1 < len(raw):
        buf = re.sub(r"\s_\s*$", " ", buf)
        i += 1
        buf += raw[i].strip()
    logical.append((start, buf))
    i += 1

code = [strip_code(t) for _, t in logical]
lnos = [n for n, _ in logical]

# ---- 프로시저 목록 ----
procdef = {}
for (ln, t) in zip(lnos, code):
    m = re.match(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", t, re.I)
    if m:
        procdef[m.group(1).lower()] = ln
defined = set(procdef)

# ---- 블록 균형 ----
issues = []
stack = []


def push(kind, ln):
    stack.append((kind, ln))


def close(kind, ln):
    for i2 in range(len(stack) - 1, -1, -1):
        if stack[i2][0] == kind:
            if i2 != len(stack) - 1:
                issues.append(f"L{ln}: '{kind}' 닫기 전 미종료 {stack[i2+1:]}")
            del stack[i2:]
            return
    issues.append(f"L{ln}: 짝 없는 '{kind}'")


for ln, t, l in zip(lnos, code, code):
    s = t.strip()
    lo = s.lower()
    if not s:
        continue
    if re.match(r"^(?:public\s+|private\s+|friend\s+)?(?:sub|function|property\s+(?:get|let|set))\s+", lo):
        push("proc", ln); continue
    if re.match(r"^end\s+(sub|function|property)\s*$", lo):
        close("proc", ln); continue
    if re.match(r"^with\b", lo):
        push("with", ln); continue
    if re.match(r"^end\s+with\b", lo):
        close("with", ln); continue
    if re.match(r"^select\s+case\b", lo):
        push("select", ln); continue
    if re.match(r"^end\s+select\b", lo):
        close("select", ln); continue
    if re.match(r"^for\b", lo):
        push("for", ln); continue
    if re.match(r"^next\b", lo):
        close("for", ln); continue
    if re.match(r"^do\b", lo):
        push("do", ln); continue
    if re.match(r"^loop\b", lo):
        close("do", ln); continue
    if re.match(r"^if\b", lo):
        rest = s[2:]
        mt = re.search(r"\bthen\b", rest, re.I)
        if mt and rest[mt.end():].strip() != "":
            pass                      # 한 줄 If
        else:
            push("if", ln)
        continue
    if re.match(r"^else\b", lo):
        continue
    if re.match(r"^end\s+if\s*$", lo):
        close("if", ln); continue
    if re.match(r"^type\s+\w", lo):
        push("type", ln); continue
    if re.match(r"^end\s+type\b", lo):
        close("type", ln); continue
    if re.match(r"^enum\b", lo):
        push("enum", ln); continue
    if re.match(r"^end\s+enum\b", lo):
        close("enum", ln); continue
    if re.match(r"^case\b", lo):
        if not any(k == "select" for k, _ in stack):
            issues.append(f"L{ln}: Select Case 밖의 Case")
        continue

for k, ln in stack:
    issues.append(f"L{ln}: 닫히지 않은 '{k}'")

print("=== 1. 블록 균형 ===")
if issues:
    for x in issues:
        print("  ✗", x)
else:
    print("  ✓ 이상 없음")

# ---- Select Case 중복 Case ----
print("\n=== 2. Select Case 중복 Case ===")
dups = []
cur = None
seen = set()
for ln, s in zip(lnos, code):
    lo = s.strip().lower()
    if re.match(r"^select\s+case\b", lo):
        cur, seen = ln, set()
    elif re.match(r"^end\s+select\b", lo):
        cur, seen = None, set()
    elif cur and re.match(r"^case\b", lo):
        key = re.sub(r"\s+", "", lo)
        if key in seen:
            dups.append((cur, ln, s.strip()))
        seen.add(key)
if dups:
    for c, l, t in dups:
        print(f"  ✗ Select(L{c}) 안 L{l} 중복: {t}")
else:
    print("  ✓ 이상 없음")

# ---- Range = Array() 크기 ----
print("\n=== 3. Range = Array() 크기 ===")
bad = 0
for idx, (ln, s) in enumerate(zip(lnos, code)):
    m = re.search(r'Range\("([A-Z])(\d+):([A-Z])(\d+)"\)\s*=\s*Array\(', s)
    if not m:
        continue
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    ci = lambda c: ord(c) - 64
    ncell = (ci(c2) - ci(c1) + 1) * (r2 - r1 + 1)
    buf = s[m.end():]
    j = idx
    while buf.count("(") > buf.count(")") and j + 1 < len(code):
        j += 1
        buf += " " + code[j]
    inner = buf[:buf.rfind(")")] if ")" in buf else buf
    depth, inq, cnt = 0, False, (1 if inner.strip() else 0)
    for ch in inner:
        if ch == '"':
            inq = not inq
        elif not inq and ch in "([":
            depth += 1
        elif not inq and ch in ")]":
            depth -= 1
        elif ch == "," and not inq and depth == 0:
            cnt += 1
    if cnt != ncell:
        bad += 1
        print(f"  ✗ L{ln}: Range {m.group(1)}{m.group(2)}:{m.group(3)}{m.group(4)} = {ncell}칸 / Array {cnt}개")
if bad == 0:
    print("  ✓ 이상 없음")

# ---- 미정의 프로시저 ----
print("\n=== 4. 미정의 프로시저 호출 ===")
builtin = set("""msgbox format left right mid trim ucase lcase len cstr cint clng cdbl
csng abs int fix sgn sqr sin cos atn rnd val str isnumeric isempty isnull isarray
instr instrrev replace split join array ubound lbound chr asc now timer iif switch
choose rg b doevents appactivate sendkeys inputbox dir createobject getobject
evaluate callbyname typename vartype round exp log tan hex oct cbool cbyte ccur
cdate cvdate cverr date dateadd datediff datepart dateserial datevalue day hour
minute month second time timeserial timevalue weekday year filedatetime
filelen getattr setattr strcomp strconv strreverse space string tab
vbcr vbcrlf vblf vbtab vbnullchar vbnullstring vbobjecterror
vbinformation vbcritical vbexclamation vbquestion vbyesno vbyes vbno vbokonly
vbok vbcancel vbretry vbignore vbabor t vbok vbdefaultbutton1 vbdefaultbutton2
xlnone xlup xldown xltoleft xltoright xlcontinuous xlthin xlmedium xlcenter
xlleft xlright xltop xlbottom xlautomatic rgb qbcolor vbnewline""".split())
calls = collections.Counter()
for ln, s in zip(lnos, code):
    for m in re.finditer(r"(?<![.\w$])([A-Za-z_][A-Za-z0-9_]*)\s*\(", s):
        nm = m.group(1)
        pre = s[:m.start()].rstrip()
        if pre.endswith("."):
            continue
        if re.search(r"\b(sub|function|property\s+\w+)\s*$", pre, re.I):
            continue
        calls[nm] += 1
missing = [(nm, n) for nm, n in calls.items()
           if nm.lower() not in defined and nm.lower() not in builtin
           and not re.match(r"^[A-Z]{1,3}$", nm)]   # 1~3자 대문자는 배열/변수로 간주
if missing:
    for nm, n in sorted(missing, key=lambda x: -x[1]):
        print(f"  ? {nm} ({n}회)")
else:
    print("  ✓ 이상 없음")

# ---- 변수 선언 여부 (Option Explicit 대비) ----
print("\n=== 5. Option Explicit 대비 미선언 식별자(휴리스틱) ===")
print("  (건너뜀 — Excel 에서 'Debug > Compile' 로 확인 권장)")

print(f"\n프로시저 {len(defined)}개 / 발견 문제 {len(issues)+len(dups)+bad+len(missing)}건")
