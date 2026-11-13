#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""VBA Option Explicit 검사기 v3
- 모듈 스코프 선언을 '파일 전체'에서 수집 (프로시저 본문 밖 전부)
- 프로시저별 파라미터 + 지역 Dim/Const/ReDim + 모듈변수 + 내장 + 호스트전역
- 사용된 식별자 중 미선언만 보고
"""
import re
import sys

PATH = sys.argv[1] if len(sys.argv) > 1 else '_src/횡단면도그리기_utf8.bas'
raw = open(PATH, encoding='utf-8').read().split('\n')

# ── 키워드 ────────────────────────────────────────────────
KW = set("""and as byref byval call case const declare dim do each else elseif empty end
enum eqv erase event exit false for friend function get global gosub goto if imp
implements in is let like lock loop lset me mod new next not nothing null on option
optional or paramarray preserve private property public put raiseevent redim rem
resume return rset select set static step stop sub then to true type typeof until
wend while with withevents xor
boolean byte currency date decimal double integer long object single string variant
any ptr err on off resume goto
""".split())

# VBA 형식/개체 형식 이름 (선언문에 타입으로 등장)
TYPES = set("""worksheet workbook range collection object variant string integer long
double single boolean byte currency date decimal any
""".split())

BUILTIN = set("""msgbox inputbox format left right mid trim lcase ucase len cstr clng
cint cdbl csng cbool cdate cbyte ccur cvdate cverr isnumeric isempty isnull isarray
isobject ismissing iserror instr instrrev replace split join array ubound lbound
abs int fix sgn sqr sin cos tan atn exp log round rnd val str strptr chr chrw asc
ascw hex oct now date time timer doevents appactivate sendkeys createobject
getobject callbyname evaluate typename vartype iif switch choose rgb qbcolor space
string strcomp strconv strreverse filter formatnumber formatcurrency
dateadd datediff datepart dateserial datevalue day hour minute month second
timeserial timevalue weekday year
sheets worksheets workbooks range cells rows columns
application thisworkbook activeworkbook activesheet activecell selection
activewindow activeselection
""".split())

HOST = set("""acad doc ms cadappname""".split())

# 정렬 상수 / 시스템 상수 (vb*, xl*, ac*)
SYSCONST = re.compile(r'^(vb|xl|ac|mso|wiz|ofc|cst)[A-Z]', re.I)

# ── 주석·문자열 제거 ──────────────────────────────────────
def clean(s):
    out, inq = [], False
    for c in s:
        if c == '"':
            inq = not inq
            out.append(' ')
        elif c == "'" and not inq:
            break
        else:
            out.append(' ' if inq else c)
    return ''.join(out)

# ── 라인 컨티뉴에이션 병합 ────────────────────────────────
logical = []   # (start_line, end_line, code_text)
i = 0
while i < len(raw):
    st = i + 1
    buf, cl = raw[i], clean(raw[i])
    while re.search(r'\s_\s*$', cl) and i + 1 < len(raw):
        buf = re.sub(r'\s_\s*$', ' ', buf)
        i += 1
        buf += raw[i].strip()
        cl = clean(buf)
    logical.append((st, i + 1, cl))
    i += 1

# ── 프로시저 경계 ────────────────────────────────────────
SIG = re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)?'
                 r'(Sub|Function|Property\s+(?:Get|Let|Set))\s+'
                 r'([A-Za-z_가-힣][\w]*)\s*(\()?', re.I)
ENDS = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.I)

procs = []
idx = 0
while idx < len(logical):
    t = logical[idx][2]
    m = SIG.match(t)
    if m and not ENDS.match(t) and not re.match(r'^\s*(End|Exit)\b', t, re.I):
        end = idx
        for j in range(idx + 1, len(logical)):
            if ENDS.match(logical[j][2]):
                end = j
                break
        procs.append(dict(name=m.group(2), kind=m.group(1),
                          sig=idx, end=end,
                          line=logical[idx][0], sigtext=logical[idx][2]))
        idx = end + 1
    else:
        idx += 1

# ── 모듈 스코프 = 프로시저 본문에 속하지 않는 모든 줄 ────
in_body = [False] * len(logical)
for p in procs:
    for j in range(p['sig'], p['end'] + 1):
        in_body[j] = True

# ── 식별자 분해 ──────────────────────────────────────────
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

ID = re.compile(r'([A-Za-z_가-힣][\w]*|\[[^\]]+\])')

PREFIXES = re.compile(
    r'^\s*(?:WithEvents|Static|Optional|ByVal|ByRef|ParamArray|Preserve|'
    r'Public|Private|Global|Dim|Friend)\s+', re.I)

def names_of(decl_text):
    """'a As X, b(0 To 3) As Y, c!' -> ['a','b','c']"""
    out = []
    for p in split_top(decl_text):
        # 접두사가 여러 개 붙을 수 있다 (예: 'Optional ByVal rot ...')
        for _ in range(8):
            np = PREFIXES.sub('', p, count=1)
            if np == p:
                break
            p = np
        m = ID.match(p.strip())
        if not m:
            continue
        nm = m.group(1).strip('[]')
        if nm.lower() in KW:
            continue
        if not nm or nm[0] == '#':
            continue
        out.append(nm)
    return out

def param_names(sigtext):
    m = re.search(r'\(', sigtext)
    if not m:
        return []
    depth, body = 0, ''
    for k in range(m.start(), len(sigtext)):
        ch = sigtext[k]
        if ch == '(':
            depth += 1
            if depth == 1:
                continue
        elif ch == ')':
            depth -= 1
            if depth == 0:
                break
        body += ch
    return names_of(body)

# ── 모듈 레벨 선언 수집 (전체 파일) ──────────────────────
module_vars = {}
module_types = set()
in_type = False
for j, (s, e, t) in enumerate(logical):
    if in_body[j]:
        continue
    st = t.strip()
    if not st:
        continue
    if re.match(r'^(?:Public\s+|Private\s+)?Type\s+', st, re.I):
        in_type = True
        mm = re.match(r'^(?:Public\s+|Private\s+)?Type\s+([\w가-힣]+)', st, re.I)
        if mm:
            module_types.add(mm.group(1).lower())
        continue
    if in_type:
        if re.match(r'^End\s+Type', st, re.I):
            in_type = False
        continue
    if re.match(r'^(?:Public\s+|Private\s+|Global\s+|Dim\s+|Friend\s+)?Declare\b', st, re.I):
        continue
    m = re.match(r'^(?:Public\s+|Private\s+|Global\s+|Friend\s+|Dim\s+)?Const\s+(.+)$', st, re.I)
    if m:
        for nm in names_of(m.group(1)):
            module_vars[nm.lower()] = nm
        continue
    m = re.match(r'^(?:Public\s+|Private\s+|Global\s+|Dim\s+|Friend\s+)(.+)$', st, re.I)
    if m:
        for nm in names_of(m.group(1)):
            module_vars[nm.lower()] = nm

proc_names = set(p['name'].lower() for p in procs)

# ── 프로시저별 검사 ─────────────────────────────────────
LABEL = re.compile(r'^([A-Za-z_가-힣][\w]*):\s*$')

report = []
for p in procs:
    local = set()
    for j in range(p['sig'] + 1, p['end']):
        st = logical[j][2].strip()
        m = re.match(r'^(?:Dim|Static|Const)\s+(.+)$', st, re.I)
        if m:
            for nm in names_of(m.group(1)):
                local.add(nm.lower())
            continue
        m = re.match(r'^ReDim\s+(?:Preserve\s+)?(.+)$', st, re.I)
        if m:
            for nm in names_of(m.group(1)):
                local.add(nm.lower())

    params = set(x.lower() for x in param_names(p['sigtext']))
    known = (local | params | set(module_vars) | module_types
             | proc_names | BUILTIN | KW | HOST)

    bad = []
    for j in range(p['sig'] + 1, p['end']):
        s, e, t = logical[j]
        st = t.strip()
        if not st:
            continue
        # 줄 레이블은 제외
        if LABEL.match(st):
            continue
        # obj.Prop / obj!Prop 제거
        w = re.sub(r'\.[A-Za-z_가-힣][\w]*', '', t)
        w = re.sub(r'![A-Za-z_가-힣][\w]*', '', w)
        # 이름있는 인자 Name:= 제거
        w = re.sub(r'([A-Za-z_가-힣][\w]*)\s*:=', '', w)
        # GoTo/GoSub 대상, On Error GoTo 라벨 제거
        w = re.sub(r'\b(GoTo|GoSub)\s+[A-Za-z_가-힣][\w]*', '', w, flags=re.I)
        # 문자열/주석 제거
        w = clean(w)

        for m in re.finditer(r'[A-Za-z_가-힣][\w]*', w):
            nm = m.group(0)
            lo = nm.lower()
            if lo in known or lo in KW:
                continue
            if SYSCONST.match(nm):
                continue
            # 순수 숫자/단위 접미사 제외
            if len(nm) == 1 and nm in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ':
                pass
            bad.append((nm, s))

    if bad:
        uniq = {}
        for nm, ln in bad:
            uniq.setdefault(nm, ln)
        report.append((p['name'], p['line'], sorted(uniq.items(), key=lambda x: x[1])))

print('프로시저 %d개 / 모듈변수 %d개 / 모듈Const 포함' % (len(procs), len(module_vars)))
print()
if not report:
    print('✓ 미선언 식별자 없음')
else:
    print('=== 미선언 의심 식별자 ===')
    tot = 0
    for nm, ln, items in report:
        tot += len(items)
        print('  [%s]  L%d' % (nm, ln))
        for a, b in items:
            print('        %-24s L%d' % (a, b))
    print()
    print('합계:', tot)
