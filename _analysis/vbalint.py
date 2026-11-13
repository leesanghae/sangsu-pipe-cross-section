# -*- coding: utf-8 -*-
"""VBA 컴파일 오류 유발 패턴 전수 린터"""
import re, sys

path = sys.argv[1] if len(sys.argv) > 1 else '_src/횡단면도그리기_utf8.bas'
raw = open(path, encoding='utf-8').read().split('\n')

def strip_code(s):
    out, inq = [], False
    for c in s:
        if c == '"': inq = not inq; out.append(c)
        elif c == "'" and not inq: break
        else: out.append(c)
    return "".join(out)

# 라인 컨티뉴에이션 병합
logical, i = [], 0
while i < len(raw):
    start, buf, cnt = i+1, raw[i], 0
    while re.search(r"\s_\s*$", strip_code(buf)) and i+1 < len(raw):
        cnt += 1
        buf = re.sub(r"\s_\s*$", " ", buf); i += 1; buf += raw[i].strip()
    logical.append((start, buf, cnt)); i += 1
code = [strip_code(t) for _, t, _ in logical]
lnos = [n for n, _, _ in logical]
cont = [c for _, _, c in logical]

problems = []

# 1) 프로시저 타입 일치 (Sub->End Sub, Function->End Function)
PROC = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?"
                  r"(Sub|Function)\s+([^\s(]+)", re.I)
stack = []
for ln, s in zip(lnos, code):
    m = PROC.match(s)
    if m:
        stack.append((m.group(1).lower(), m.group(2), ln)); continue
    m2 = re.match(r"^\s*End\s+(Sub|Function)\s*$", s, re.I)
    if m2:
        if not stack:
            problems.append((ln, "짝 없는 End %s" % m2.group(1))); continue
        k, nm, dl = stack.pop()
        if k != m2.group(1).lower():
            problems.append((ln, "L%d %s %s -> 'End %s' (타입 불일치!)" % (dl, k, nm, m2.group(1))))
for k, nm, dl in stack:
    problems.append((dl, "%s %s 가 닫히지 않음" % (k, nm)))
print("1) Sub/Function - End 타입 일치: %s" % ("✓" if not problems else "✗ %d건" % len(problems)))

# 2) Exit Sub/Function 이 맞는 프로시저 안인지
p2 = []
st = []
for ln, s in zip(lnos, code):
    m = PROC.match(s)
    if m: st.append(m.group(1).lower()); continue
    if re.match(r"^\s*End\s+(Sub|Function)\s*$", s, re.I):
        if st: st.pop()
        continue
    m3 = re.match(r"^\s*Exit\s+(Sub|Function)\s*$", s, re.I)
    if m3 and st and st[-1] != m3.group(1).lower():
        p2.append((ln, "Exit %s 인데 여기는 %s" % (m3.group(1), st[-1])))
for ln, t in p2: problems.append((ln, t))
print("2) Exit Sub/Function 정합: %s" % ("✓" if not p2 else "✗ %d건" % len(p2)))

# 3) Optional 파라미터 뒤에 필수 파라미터 (컴파일 오류)
p3 = []
for ln, s, c in zip(lnos, code, cont):
    if not PROC.match(s): continue
    # 파라미터 목록 추출 (병합된 논리행 기준)
    if '(' not in s: continue
    ps = s[s.index('('):]
    ps = ps[:ps.rfind(')')] if ')' in ps else ps
    parts, depth, cur = [], 0, ''
    for ch in ps:
        if ch in '([': depth += 1; cur += ch
        elif ch in ')]': depth -= 1; cur += ch
        elif ch == ',' and depth == 0: parts.append(cur); cur = ''
        else: cur += ch
    if cur.strip(): parts.append(cur)
    seen_opt = False
    for p in parts:
        p = p.strip()
        if not p: continue
        is_opt = re.match(r"^(?:ByVal\s+|ByRef\s+)?Optional\b", p, re.I) is not None
        if is_opt: seen_opt = True
        elif seen_opt:
            p3.append((ln, "Optional 뒤에 필수 파라미터: %s" % p[:40]))
for ln, t in p3: problems.append((ln, t))
print("3) Optional 뒤 필수 파라미터: %s" % ("✓" if not p3 else "✗ %d건" % len(p3)))

# 4) 파라미터 이름 중복
p4 = []
for ln, s in zip(lnos, code):
    if not PROC.match(s) or '(' not in s: continue
    ps = s[s.index('('):]
    ps = ps[:ps.rfind(')')] if ')' in ps else ps
    names = re.findall(r"(?:ByVal|ByRef)?\s*(?:Optional\s+)?([A-Za-z_]\w*)\s*(?:\(\))?\s*(?:As\s+[\w\.]+)?\s*(?:=|,|$)", ps)
    seen = set()
    for n in names:
        nl = n.lower()
        if nl in ('byval','byref','optional','as'): continue
        if nl in seen: p4.append((ln, "파라미터 이름 중복: %s" % n))
        seen.add(nl)
for ln, t in p4: problems.append((ln, t))
print("4) 파라미터 이름 중복: %s" % ("✓" if not p4 else "✗ %d건" % len(p4)))

# 5) 라인 컨티뉴에이션 24개 초과
p5 = [(ln, "컨티뉴에이션 %d개 (VBA 한도 24)" % c) for ln, c in zip(lnos, cont) if c > 24]
for ln, t in p5: problems.append((ln, t))
print("5) 라인 컨티뉴에이션 ≤24: %s" % ("✓" if not p5 else "✗ %d건" % len(p5)))

# 6) Option Explicit 위치
p6 = []
idx = [i for i, s in enumerate(code) if re.match(r"^\s*Option\s+Explicit\s*$", s, re.I)]
if not idx: p6.append((1, "Option Explicit 없음"))
else:
    firstp = next((ln for ln, s in zip(lnos, code) if PROC.match(s)), 10**9)
    if lnos[idx[0]] > firstp: p6.append((lnos[idx[0]], "Option Explicit 가 프로시저 뒤에 있음"))
for ln, t in p6: problems.append((ln, t))
print("6) Option Explicit 최상단: %s" % ("✓" if not p6 else "✗ %d건" % len(p6)))

# 7) 프로시저 이름 중복 (모듈 내)
p7 = []
seen = {}
for ln, s in zip(lnos, code):
    m = PROC.match(s)
    if m:
        nm = m.group(2).lower()
        if nm in seen: p7.append((ln, "프로시저 중복 정의 '%s' (앞: L%d)" % (m.group(2), seen[nm])))
        else: seen[nm] = ln
for ln, t in p7: problems.append((ln, t))
print("7) 프로시저 이름 중복: %s" % ("✓" if not p7 else "✗ %d건" % len(p7)))

print()
if problems:
    print("*** 문제 %d건 ***" % len(problems))
    for ln, t in sorted(problems)[:60]:
        print("   L%-5d %s" % (ln, t))
else:
    print("*** 컴파일 유발 패턴 없음 ***")
