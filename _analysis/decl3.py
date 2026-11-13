#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""고정밀 미선언 변수 검사
- '반드시 선언되어야 하는 위치'만 검사: 대입 LHS, For 카운터, For Each, With 대상, ReDim
- 모듈 스코프 선언은 파일 전체에서 수집 (프로시저 본문 제외)
- 파라미터는 다중 접두사(Optional ByVal ...) 처리
- CP949 파일의 '?' 치환 여부도 검사
"""
import re, sys

UTF8 = '_src/횡단면도그리기_utf8.bas'
ANSI = '_src/횡단면도그리기.bas'

def clean(s):
    out, inq = [], False
    for c in s:
        if c == '"':
            inq = not inq; out.append(' ')
        elif c == "'" and not inq:
            break
        else:
            out.append(' ' if inq else c)
    return ''.join(out)

def merge(raw):
    logical, i = [], 0
    while i < len(raw):
        st, buf, cl = i + 1, raw[i], clean(raw[i])
        while re.search(r'\s_\s*$', cl) and i + 1 < len(raw):
            buf = re.sub(r'\s_\s*$', ' ', buf)
            i += 1
            buf += raw[i].strip(); cl = clean(buf)
        logical.append((st, i + 1, cl)); i += 1
    return logical

KW = set("""and as byref byval call case const dim do each else elseif empty end enum
eqv erase event exit false for friend function get global gosub goto if imp in is let
like loop lset me mod new next not nothing null on option optional or paramarray
preserve private property public put raiseevent redim rem resume return rset select
set static step stop sub then to true type until wend while with withevents xor
boolean byte currency date decimal double integer long object single string variant
""".split())

PREFIX = re.compile(r'^\s*(?:WithEvents|Static|Optional|ByVal|ByRef|ParamArray|'
                    r'Preserve|Public|Private|Global|Dim|Friend)\s+', re.I)
ID = re.compile(r'([A-Za-z_가-힣][\w]*)')

def split_top(s):
    parts, depth, cur = [], 0, ''
    for ch in s:
        if ch in '([': depth += 1
        elif ch in ')]': depth -= 1
        if ch == ',' and depth == 0:
            parts.append(cur); cur = ''
        else: cur += ch
    parts.append(cur)
    return parts

def names_of(text):
    out = []
    for p in split_top(text):
        for _ in range(8):
            np = PREFIX.sub('', p, count=1)
            if np == p: break
            p = np
        m = ID.match(p.strip())
        if m and m.group(1).lower() not in KW:
            out.append(m.group(1))
    return out

SIG = re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)?'
                 r'(Sub|Function|Property\s+(?:Get|Let|Set))\s+'
                 r'([A-Za-z_가-힣][\w]*)\s*(\()?', re.I)
ENDS = re.compile(r'^\s*End\s+(Sub|Function|Property)\b', re.I)

def analyze(path, label):
    raw = open(path, encoding='utf-8', errors='replace').read().split('\n')
    logical = merge(raw)

    procs, i = [], 0
    while i < len(logical):
        t = logical[i][2]
        m = SIG.match(t)
        if m and not ENDS.match(t) and not re.match(r'^\s*(End|Exit)\b', t, re.I):
            end = len(logical) - 1
            for j in range(i + 1, len(logical)):
                if ENDS.match(logical[j][2]):
                    end = j; break
            procs.append(dict(name=m.group(2), sig=i, end=end,
                              line=logical[i][0], sigtext=t))
            i = end + 1
        else:
            i += 1

    in_body = [False] * len(logical)
    for p in procs:
        for j in range(p['sig'], p['end'] + 1):
            in_body[j] = True

    # 모듈 스코프 선언 (파일 전체)
    module = set()
    in_type = False
    for j, (s, e, t) in enumerate(logical):
        if in_body[j]: continue
        st = t.strip()
        if not st: continue
        if re.match(r'^(?:Public\s+|Private\s+)?Type\s+', st, re.I):
            in_type = True; continue
        if in_type:
            if re.match(r'^End\s+Type', st, re.I): in_type = False
            continue
        if re.match(r'^(?:Public|Private|Global|Friend)\s+Declare\b', st, re.I):
            continue
        m = re.match(r'^(?:(?:Public|Private|Global|Friend|Dim)\s+)?Const\s+(.+)$', st, re.I)
        if m:
            for nm in names_of(m.group(1)): module.add(nm.lower())
            continue
        m = re.match(r'^(?:Public|Private|Global|Dim|Friend)\s+(.+)$', st, re.I)
        if m:
            for nm in names_of(m.group(1)): module.add(nm.lower())

    procnames = set(p['name'].lower() for p in procs)
    BUILTIN = set("""err msgbox format left right mid trim lcase ucase len cstr clng
    cint cdbl csng cbool cdate isnumeric isempty isnull isarray isobject iserror
    instr instrrev replace split join array ubound lbound abs int fix sgn sqr sin
    cos tan atn exp log round rnd val str chr asc hex oct now date time timer
    doevents appactivate sendkeys createobject getobject callbyname evaluate
    typename vartype iif switch choose rgb space string strcomp strconv
    thisworkbook application worksheets sheets workbooks""".split())
    HOST = set('acad doc ms cadappname'.split())

    problems = []
    for p in procs:
        local = set()
        for j in range(p['sig'] + 1, p['end']):
            st = logical[j][2].strip()
            m = re.match(r'^(?:Dim|Static|Const)\s+(.+)$', st, re.I)
            if m:
                for nm in names_of(m.group(1)): local.add(nm.lower()); continue
            m = re.match(r'^ReDim\s+(?:Preserve\s+)?(.+)$', st, re.I)
            if m:
                for nm in names_of(m.group(1)): local.add(nm.lower())

        # 파라미터
        params = set()
        sig = p['sigtext']
        if '(' in sig:
            depth, body = 0, ''
            for k in range(sig.index('('), len(sig)):
                ch = sig[k]
                if ch == '(':
                    depth += 1
                    if depth == 1: continue
                elif ch == ')':
                    depth -= 1
                    if depth == 0: break
                body += ch
            for nm in names_of(body): params.add(nm.lower())

        known = local | params | module | procnames | BUILTIN | KW | HOST

        for j in range(p['sig'] + 1, p['end']):
            s0, e0, t = logical[j]
            st = t.strip()
            if not st: continue
            # obj.Prop / obj!Prop 제거
            w = re.sub(r'[.!][A-Za-z_가-힣][\w]*', '', t)
            w = re.sub(r'([A-Za-z_가-힣][\w]*)\s*:=', '', w)   # named arg
            w = re.sub(r'\b(GoTo|GoSub|On\s+Error\s+GoTo)\s+[A-Za-z_가-힣][\w]*', '', w, flags=re.I)

            cands = []
            # For X = / For Each X In
            m = re.match(r'^\s*For\s+Each\s+([A-Za-z_가-힣][\w]*)\s+In\b', w, re.I)
            if m: cands.append((m.group(1), 'ForEach'))
            m = re.match(r'^\s*For\s+([A-Za-z_가-힣][\w]*)\s*=', w, re.I)
            if m: cands.append((m.group(1), 'For'))
            # With X
            m = re.match(r'^\s*With\s+([A-Za-z_가-힣][\w]*)\s*$', w, re.I)
            if m: cands.append((m.group(1), 'With'))
            # ReDim X(
            m = re.match(r'^\s*ReDim\s+(?:Preserve\s+)?([A-Za-z_가-힣][\w]*)', w, re.I)
            if m: cands.append((m.group(1), 'ReDim'))
            # LHS 대입:  [Set] X = ...   (X 는 식별자 하나 또는 X(...))
            m = re.match(r'^\s*(?:Set\s+)?([A-Za-z_가-힣][\w]*)\s*(\([^)]*\))?\s*=(?!=)', w)
            if m and not re.match(r'^\s*(If|ElseIf|Else|While|Until|Loop|Case|End|'
                                  r'On|Function|Sub|Property)\b', w, re.I):
                cands.append((m.group(1), 'assign'))

            for nm, why in cands:
                lo = nm.lower()
                if lo in known: continue
                problems.append((p['name'], s0, nm, why))

    print('=== %s ===' % label)
    print('프로시저 %d / 모듈선언 %d' % (len(procs), len(module)))
    if problems:
        seen = set()
        for pn, ln, nm, why in problems:
            k = (pn, nm)
            if k in seen: continue
            seen.add(k)
            print('  ✗ %-24s L%-6d %-22s (%s)' % (pn, ln, nm, why))
    else:
        print('  ✓ 미선언 변수 없음')
    print()
    return problems

p1 = analyze(UTF8, 'UTF-8 원본')
p2 = analyze(ANSI, 'CP949 사본 (실제 임포트본)')

# CP949 '?' 치환 검사 (코드 줄만)
print('=== CP949 변환 손상 검사 ===')
raw = open(ANSI, 'rb').read()
txt = raw.decode('cp949', errors='replace')
bad_lines = []
for i, line in enumerate(txt.split('\n'), 1):
    code = clean(line)
    if '?' in code:
        bad_lines.append((i, line.strip()[:90]))
if bad_lines:
    print('  ✗ 코드 줄에 "?" 존재 (인코딩 치환 의심):', len(bad_lines))
    for ln, t in bad_lines[:10]: print('     L%d: %s' % (ln, t))
else:
    print('  ✓ 코드 줄에 "?" 없음')
print('  주석 내 "?" 개수:', sum(clean(l).count('?') for l in []))
full = txt.count('?')
print('  파일 전체 "?" 개수:', full, '(주석의 특수문자 치환분)')
