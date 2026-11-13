#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VBA Option Explicit 검사기 — 선언되지 않은 변수를 프로시저별로 찾아낸다."""
import re
import sys

PATH = '_src/횡단면도그리기_utf8.bas'
raw = open(PATH, encoding='utf-8').read().split('\n')

KW = set("""and as byref byval call case const declare dim do each else elseif empty end
enum eqv err erase event exit false for friend function get global gosub goto if imp
implements in is let like lock loop lset me mod new next not nothing null on option
optional or paramarray preserve private property public put raiseevent redim rem
resume return rset select set static step stop sub then to true type typeof until
wend while with withevents xor
boolean byte currency date decimal double integer long object single string variant
any ptr
debug print
vbcrlf vblf vbcr vbtab vbnewline vbnullstring vbnullchar
""".split())

BUILTIN = set("""msgbox inputbox format left right mid trim lcase ucase len cstr clng
cint cdbl csng cbool cdate cbyte ccur cvdate cverr isnumeric isempty isnull isarray
isobject ismissing iserror instr instrrev replace split join array ubound lbound
abs int fix sgn sqr sin cos tan atn exp log round rnd val str chr asc hex oct
now date time timer doevents appactivate sendkeys createobject getobject callbyname
evaluate typename vartype iif switch choose rgb qbcolor space string strcomp
strconv strreverse filter weekdaydat e dateadd datediff datepart dateserial
datevalue day hour minute month second timeserial timevalue year
sheets worksheets workbooks range cells activecell activesheet activeworkbook
thisworkbook application selection
""".split())

# ---------- 주석/문자열 제거 ----------
def clean(s):
    out = []
    inq = False
    i = 0
    while i < len(s):
        c = s[i]
        if c == '"':
            inq = not inq
            out.append(' ')
        elif c == "'" and not inq:
            break
        else:
            out.append(' ' if inq else c)
        i += 1
    return ''.join(out)

# ---------- 라인 컨티뉴에이션 병합 ----------
logical = []   # (start_ln, end_ln, clean_text)
i = 0
while i < len(raw):
    st = i + 1
    buf = raw[i]
    cl = clean(buf)
    while re.search(r'\s_\s*$', cl) and i + 1 < len(raw):
        buf = re.sub(r'\s_\s*$', ' ', buf)
        i += 1
        buf += raw[i].strip()
        cl = clean(buf)
    logical.append((st, i + 1, cl))
    i += 1

# ---------- 프로시저 경계 ----------
SIG = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)?'
                 r'(Sub|Function|Property\s+(?:Get|Let|Set))\s+'
                 r'([A-Za-z_가-힣\w]*)\s*(\(.*)?$', re.I)

procs = []       # dict(name, kind, sig_idx, body_idx, end_idx, params)
for idx, (s, e, t) in enumerate(logical):
    m = SIG.match(t)
    if m and not re.match(r'^\s*End\s', t, re.I):
        # 같은 logical 안에 End Sub/Function 이 있으면 한 줄 프로시저
        procs.append(dict(name=m.group(2), kind=m.group(1), sig=idx, end=None,
                          sig_line=s, sig_end_line=e))

for p in procs:
    for j in range(p['sig'] + 1, len(logical)):
        t = logical[j][2].strip()
        if re.match(r'^End\s+(Sub|Function|Property)\b', t, re.I):
            p['end'] = j
            break
    if p['end'] is None:
        p['end'] = len(logical) - 1

first_proc_line = procs[0]['sig'] if procs else len(logical)

# ---------- 이름 추출 (Dim a As X, b(0 To 3) As Y) ----------
def split_top(s):
    parts, depth, cur = [], 0, ''
    for ch in s:
        if ch in '([':
            depth += 1
        elif ch in ')]':
            depth -= 1
        if ch == ',' and depth == 0:
            parts.append(cur); cur = ''
        else:
            cur += ch
    parts.append(cur)
    return parts

IDENT = re.compile(r'([A-Za-z_가-힣][\w]*)')

def decl_names(rest):
    rest = re.sub(r'^\s*Const\s+', '', rest, flags=re.I)
    out = []
    for p in split_top(rest):
        p = re.sub(r'^\s*(?:WithEvents\s+|Static\s+)', '', p, flags=re.I)
        m = IDENT.match(p.strip())
        if not m:
            continue
        nm = m.group(1)
        if nm.lower() in KW:
            continue
        out.append(nm)
    return out

def param_names(sig):
    m = re.search(r'\(', sig)
    if not m:
        return []
    depth, start = 0, m.start()
    body = ''
    for k in range(start, len(sig)):
        if sig[k] == '(':
            depth += 1
            if depth == 1:
                continue
        elif sig[k] == ')':
            depth -= 1
            if depth == 0:
                break
        body += sig[k]
    out = []
    for p in split_top(body):
        p = re.sub(r'^\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?',
                   '', p, flags=re.I)
        m2 = IDENT.match(p.strip())
        if m2 and m2.group(1).lower() not in KW:
            out.append(m2.group(1))
    return out

# ---------- 모듈 레벨 선언 ----------
module_vars = {}       # lower -> original
module_types = set()
in_type = False
for idx in range(0, first_proc_line):
    t = logical[idx][2]
    s = t.strip()
    if re.match(r'^(Private|Public)?\s*Type\s+\w+', s, re.I):
        in_type = True
        module_types.add(IDENT.search(re.sub(r'^\S*\s*', '', s, count=1)).group(1).lower())
        continue
    if in_type:
        if re.match(r'^End\s+Type', s, re.I):
            in_type = False
        continue
    m = re.match(r'^\s*(?:(?:Private|Public|Global|Dim|Friend)\s+)?'
                 r'(?:(?:WithEvents|Static)\s+)?(.+)$', s, re.I)
    if re.match(r'^\s*(Private|Public|Global|Dim|Friend)\s+', s, re.I) or \
       re.match(r'^\s*(Private|Public)?\s*Const\s+', s, re.I):
        body = re.sub(r'^\s*(Private|Public|Global|Dim|Friend)\s+', '', s, flags=re.I)
        for nm in decl_names(body):
            module_vars[nm.lower()] = nm

proc_names = set(p['name'].lower() for p in procs)

# VBA / Excel / 호스트 전역 개체 (선언 불필요)
EXTRA_GLOBAL = set("""application thisworkbook activeworkbook activesheet
activecell selection activewindow activeselection
worksheets workbooks sheets charts names
ms doc acad cadappName
""".split())

# ---------- 프로시저별 검사 ----------
report = []
for p in procs:
    decl = set()
    for j in range(p['sig'] + 1, p['end']):
        t = logical[j][2].strip()
        m = re.match(r'^(Dim|Static|Const|ReDim)\s+(.+)$', t, re.I)
        if m:
            body = re.sub(r'^\s*(ReDim|Preserve)\s+', '', m.group(2), flags=re.I)
            for nm in decl_names(body):
                decl.add(nm.lower())
            continue
        m = re.match(r'^ReDim\s+Preserve\s+(.+)$', t, re.I)
        if m:
            for nm in decl_names(m.group(1)):
                decl.add(nm.lower())

    params = set(n.lower() for n in param_names(logical[p['sig']][2]))
    # 함수/프로시저명은 VBA 에서 반환값 대입 대상이므로 '선언된 것'으로 취급
    known = (decl | params | set(module_vars) | module_types
             | proc_names | EXTRA_GLOBAL)

    missing = []
    for j in range(p['sig'] + 1, p['end']):
        t = logical[j][2].strip()
        if not t:
            continue
        s = re.sub(r'\.s*[A-Za-z_가-힣]\w*', '', t)      # obj.Prop 제거
        s = re.sub(r'!s*[A-Za-z_가-힣]\w*', '', s)        # obj!Prop 제거
        # LHS 대입
        m = re.match(r'^(?:Set\s+)?([A-Za-z_가-힣]\w*)\s*=(?!=)', s, re.I)
        if m and not re.match(r'^(If|ElseIf|For|Do|Loop|While|Until|Case|End)\b', s, re.I):
            n = m.group(1).lower()
            if n not in known:
                missing.append((n, logical[j][0], 'assign'))
        # For x =
        m = re.match(r'^For\s+([A-Za-z_가-힣]\w*)\s*=', s, re.I)
        if m and m.group(1).lower() not in known:
            missing.append((m.group(1).lower(), logical[j][0], 'for'))
        # For Each x In
        m = re.match(r'^For\s+Each\s+([A-Za-z_가-힣]\w*)\s+In\b', s, re.I)
        if m and m.group(1).lower() not in known:
            missing.append((m.group(1).lower(), logical[j][0], 'foreach'))
        # With
        m = re.match(r'^With\s+([A-Za-z_가-힣]\w*)\s*$', s, re.I)
        if m and m.group(1).lower() not in known and m.group(1).lower() not in KW:
            missing.append((m.group(1).lower(), logical[j][0], 'with'))

    if missing:
        seen = []
        for n, ln, kind in missing:
            if (n, kind) not in seen:
                seen.append((n, kind))
        report.append((p['name'], p['sig_line'], seen))

# ---------- 출력 ----------
print('총 프로시저:', len(procs))
print('모듈 레벨 변수/상수:', len(module_vars))
print()
if not report:
    print('✓ 선언 누락 없음')
else:
    total = 0
    print('=== 선언 누락 (Option Explicit 오류 원인) ===')
    for nm, ln, items in report:
        total += len(items)
        print(f'  L{ln:<5} {nm:<28} ' + ', '.join(f'{a}({b})' for a, b in items))
    print()
    print('총 누락:', total)
