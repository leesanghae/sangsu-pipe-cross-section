# -*- coding: utf-8 -*-
"""
횡단면도 작도규칙 회귀검증 스펙 (Regression Spec)
================================================
DXF(정답도면, 655단면) 전체를 상대로 역설계한 작도규칙을 하나씩 검증한다.
불일치 0 = 그 규칙은 "확정". 이후 VBA 매크로의 인수인계 기준이 된다.

[중요] 셀 원점은 X0+13.0 의 근사값이 아니라, 각 단면의 '관로06' INSERT
       (= 셀 원점 + (5.0, 0)) 를 앵커로 역산해 쓴다. 오차 0.0005 수준.
"""
import math, collections
import calc

F = '/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt'
import ezdxf
doc = ezdxf.readfile(F)
msp = doc.modelspace()
rows, _ = calc.load()
N = len(rows)

# ---------------- 앵커(셀 원점) 역산 ----------------
# 논리 순서: 시트 → 열(4개) → 행(5개)
PITCH = dict(SHEET=84.1, X=18.0, Y=9.5)
N_COL, N_ROW, PER_SHEET = 4, 5, 20

_tab = [(e.dxf.insert[0], e.dxf.insert[1], {a.dxf.tag: a.dxf.text for a in e.attribs})
        for e in msp if e.dxftype() == 'INSERT' and e.dxf.name == '관로06']
BX = min(t[0] for t in _tab)          # 관로06 x = 셀원점 + 5.0
BY = min(t[1] for t in _tab)
def _key(t):
    sh = round((t[0] - BX) / PITCH['SHEET'])
    r = (t[0] - BX) - sh * PITCH['SHEET']
    return (sh, round(r / PITCH['X']), round((t[1] - BY) / PITCH['Y']))
_tab.sort(key=_key)
ANCHOR = [(t[0] - 5.0, t[1]) for t in _tab]        # 셀 원점(단면 기준점)
ATTR = [t[2] for t in _tab]

# ---------------- 엔티티 색인 (CC-BORDER 포함 전 레이어) ----------------
def xy(v):
    """Vec3 -> (x, y).  Vec3는 슬라이싱을 지원하지 않으므로 반드시 인덱스로."""
    return (float(v[0]), float(v[1]))

def pts_of(e):
    t = e.dxftype()
    try:
        if t == 'LINE':      return [xy(e.dxf.start), xy(e.dxf.end)]
        if t == 'LWPOLYLINE':return [(p[0], p[1]) for p in e.get_points('xy')]
        if t in ('TEXT', 'INSERT'): return [xy(e.dxf.insert)]
        if t == 'CIRCLE':    return [xy(e.dxf.center)]
        if t == 'HATCH':
            for p in e.paths:
                if getattr(p, 'vertices', None): return [(p.vertices[0][0], p.vertices[0][1])]
                if getattr(p, 'edges', None):
                    ed = p.edges[0]
                    c = getattr(ed, 'start', None) or getattr(ed, 'center', None)
                    if c: return [(c[0], c[1])]
    except Exception:
        pass
    return []

IDX = collections.defaultdict(list)     # layer -> [(x, y, entity)]
for e in msp:
    for p in pts_of(e):
        IDX[e.dxf.layer].append((p[0], p[1], e))
        break                            # 대표점 1개만

def near(layer, k, x0=-8.0, x1=10.0, y0=-5.6, y1=3.9):
    cx, gy = ANCHOR[k]
    return [(x, y, e) for (x, y, e) in IDX.get(layer, [])
            if cx + x0 <= x <= cx + x1 and gy + y0 <= y <= gy + y1]

# ---------------- 규칙 테이블 ----------------
def tv_of(code):
    return sum(t for _, t, r in calc.PAVE.get(code[2:4], []) if r == 'CUT')
def has_pave(code):
    return bool(calc.PAVE.get(code[2:4], []))
def bw_of(dia):
    return 0.0 if dia == 160 else 0.20      # 역설계 결과

R = collections.OrderedDict()
def chk(name, n, bad, memo=''):
    R[name] = (n, bad, memo)

TOL = 8e-3

# 원본 데이터 자체의 이상(異常) 지점 — 규칙 위반이 아니라 입력 오류.
#   k=356 : 토적표는 DIA=160 인데 SDB(H)는 DIA=80.
#           도면은 관원 r=0.04(=D80)과 토피 1.24(=H-0.08)로 그려져 SDB(H)를 따랐고,
#           표 블록의 '토피'만 1.24 로 남아 토적표 기준(1.16)과 어긋난다.
#           → 규칙은 정상. 데이터가 어긋난 지점이므로 예외로 기록한다.
EXCEPT = {356: 'SDB(H) D80 vs 토적표 D160 (중복 측점 입력오류)'}
def is_exc(k):
    return k in EXCEPT

# 1) 중심틱
chk('중심틱 CC-CENT 1개 (y ±0.04)', N,
    sum(1 for k in range(N) if len(near('CC-CENT', k)) != 1))

# 2) 지반선
chk('지반선 CC-GRND 2개 (±4.0, y=0)', N,
    sum(1 for k in range(N) if len(near('CC-GRND', k)) != 2))

# 3) 관원
b = 0; skipped = 0
for k in range(N):
    if is_exc(k):
        skipped += 1; continue
    t = rows[k]; D = t['dia'] / 1000.0; H = t['gh'] - t['inv']
    c = [e for (x, y, e) in near('CC-PIPE', k) if e.dxftype() == 'CIRCLE']
    if len(c) != 1: b += 1; continue
    cx, gy = ANCHOR[k]
    if abs(c[0].dxf.center[0] - cx) > TOL or \
       abs(c[0].dxf.center[1] - (gy - (H - D / 2))) > TOL or \
       abs(c[0].dxf.radius - D / 2) > TOL: b += 1
chk("관원 c=(0,-(H-D/2)) r=D/2 (CC-PIPE)", N, b, f'예외 {skipped}건 제외')

# 4) 표 블록
b = 0
for k in range(N):
    ins = [e for (x, y, e) in near('CC-TABL', k)
           if e.dxftype() == 'INSERT' and e.dxf.name == '관로06']
    cx, gy = ANCHOR[k]
    if len(ins) != 1 or abs(ins[0].dxf.insert[0] - (cx + 5.0)) > TOL \
       or abs(ins[0].dxf.insert[1] - gy) > TOL or abs(ins[0].dxf.xscale - 0.1) > 1e-6:
        b += 1
chk('관로06 : (+5.0,0) scale 0.1, 단면당 1개', N, b)

# 5) 굴착 폴리곤
b = 0; n_ex = 0; skipped = 0
for k in range(N):
    t = rows[k]
    if is_exc(k):
        if t['exc']: n_ex += 1; skipped += 1
        continue
    p = [e for (x, y, e) in near('CC-EXCV', k) if e.dxftype() == 'LWPOLYLINE']
    if not t['exc']:
        if p: b += 1
        continue
    n_ex += 1
    if len(p) != 1: b += 1; continue
    cx, gy = ANCHOR[k]
    pts = [(q[0] - cx, q[1] - gy) for q in p[0].get_points('xy')]
    H = t['gh'] - t['inv']; D = t['dia'] / 1000.0
    b0 = D + 2 * bw_of(t['dia']); tv = tv_of(t['code'])
    hb, ht = b0 / 2, (b0 + 2 * 0.3 * (H - tv)) / 2
    ys = [q[1] for q in pts]; xs = [abs(q[0]) for q in pts]
    # 주의: 원본 도면은 closed 플래그 False + 시작점=끝점 중복(14점)으로 닫는다.
    ok = (abs(min(ys) + H) < TOL and abs(max(ys)) < TOL
          and abs(max(xs) - ht) < TOL and abs(min(x for x in xs if x > 1e-6) - hb) < TOL
          and any(abs(q[1] + tv) < TOL and abs(abs(q[0]) - ht) < TOL for q in pts)
          and not p[0].closed
          and abs(pts[0][0] - pts[-1][0]) < 1e-9 and abs(pts[0][1] - pts[-1][1]) < 1e-9)
    if not ok: b += 1
chk('굴착 폴리곤 (바닥 b0/2, 상단 w(tv)/2@-tv, 수직벽, 시작=끝 중복)', n_ex, b,
    f'예외 {skipped}건 제외')

# 6) 해치
chk('해치 유무 = 되메우기(관주위) ≥ 0.02', N,
    sum(1 for k in range(N)
        if bool(near('CC-PATT', k)) != bool(rows[k]['exc'] and (rows[k]['ar'] or 0) >= 0.02)))

# 7) 고도축
b = 0
for k in range(N):
    t = rows[k]; cx, gy = ANCHOR[k]
    pads = [(e.dxf.start[1] - gy, e.dxf.end[1] - gy)
            for (x, y, e) in near('CC-DIML', k)
            if e.dxftype() == 'LINE' and abs(e.dxf.start[0] - e.dxf.end[0]) < 1e-6
            and e.dxf.start[0] < cx - 4.2]
    if not pads: b += 1; continue
    pb = math.floor(t['inv']) - 1
    pt = pb + 3 + math.floor(t['gh'] - t['inv'])
    if (round(t['gh'] + min(p[0] for p in pads)), round(t['gh'] + max(p[1] for p in pads))) != (pb, pt):
        b += 1
chk('고도축 bot=floor(INV)-1, top=bot+3+floor(H)', N, b)

# 8) 상세치수
chk('상세치수 CC-LEAD(28개) = 첫단면(k%20==0) & 굴착', N,
    sum(1 for k in range(N)
        if (len(near('CC-LEAD', k)) == 28) != bool(k % 20 == 0 and rows[k]['exc'])))

# 9) 포장 리더  ★스펙이 찾아낸 정련된 규칙
b = 0
for k in range(N):
    t = rows[k]
    first = (k % 20 == 0)
    chg = (k > 0 and t['code'][2:4] != rows[k - 1]['code'][2:4])
    want = (first or chg) and has_pave(t['code'])
    if bool(near('CC-LEAL', k)) != want: b += 1
chk('포장리더 CC-LEAL = (첫단면 or 포장변경) and 포장층 있음', N, b)

# 10) 라인명 라벨
b = 0
for k in range(N):
    got = any(e.dxftype() == 'TEXT' and e.dxf.text == 'SDB' for (x, y, e) in near('CC-TEXT', k))
    if got != (k % 20 == 0): b += 1
chk("라인명 라벨 'SDB' = 첫단면만", N, b)

# 11) 도곽
SHEETS = (N + PER_SHEET - 1) // PER_SHEET
bx = [x for (x, y, e) in IDX.get('CC-BORDER', []) if e.dxftype() == 'INSERT']
chk('도곽 CC-BORDER 시트당 1개 (84.1 피치)', SHEETS,
    sum(1 for s in range(SHEETS) if sum(1 for v in bx if abs(v - (bx[0] + s * PITCH['SHEET'])) < 1e-3) != 1))

# 12) 도면번호 속성
tit = sorted([(x, y, {a.dxf.tag: a.dxf.text for a in e.attribs})
              for (x, y, e) in IDX.get('CC-Title', []) if e.dxftype() == 'INSERT'])
b = 0
for i, (x, y, a) in enumerate(tit):
    exp = {'도면타이틀': f'횡 단 면 도({i+1})', '축척H': 'S = 1 : 100',
           '도면명': f'횡 단 면 도({i+1})', '축척A1': '1 : 100', '도면번호': f'C-{50+i}'}
    if any(a.get(q) != v for q, v in exp.items()): b += 1
chk('도면번호 속성 (도면명/축척H/축척A1/도번 C-50~82)', len(tit), b)

# 13) 표 블록 15개 속성 == 토적표
TAGMAP = [('측점', 'sta'), ('지반고', 'gh'), ('관저고', 'inv'), ('토피', None),
          ('터파기(육상)합계', 'exc'), ('터파기(용수)합계', 'wet'),
          ('되메우기(관주위)', 'ar'), ('되메우기(관상단)', 'ab'),
          ('ASP포장', 'asp'), ('CONC포장', 'con'), ('ASC포장', 'asc'),
          ('보도블럭포장', 'bodo'), ('되메우기(모래)', 'sand'),
          ('보조기층', 'sub'), ('동상방지층', 'frost')]
def fmt(v):
    """도면 표기 규칙: 0(또는 미입력)은 빈 문자열, 그 외는 소수 2자리.
    → 관로06 속성칸에 '0.00'이 찍힌 사례는 655×15 = 0건으로 검증됨."""
    if v is None or abs(v) < 0.005:
        return ''
    return f'{v:.2f}'
b = 0; checked = 0; miss = []; skipped = 0
for k in range(N):
    if is_exc(k):
        skipped += len(TAGMAP); continue
    a = ATTR[k]; t = rows[k]
    for tag, key in TAGMAP:
        got = a.get(tag)
        if key == 'sta':
            exp = str(t['sta'])
        elif key is None:                      # 토피 = Rnd2(H - D)
            exp = fmt(calc.rnd2(t['gh'] - t['inv'] - t['dia'] / 1000.0))
        else:
            exp = fmt(t[key])
        checked += 1
        if got != exp:
            b += 1
            if len(miss) < 5: miss.append((k, t['sta'], tag, got, exp))
chk('관로06 15속성 == 토적표 값 (0→빈칸 규칙 포함)', checked, b,
    '; '.join(f'k={m[0]}:{m[2]} DXF={m[3]!r}≠{m[4]!r}' for m in miss))

# 14) 셀 순서 == 토적표 행 순서
chk('셀 배치 순서 == 토적표 행 순서 (시트→열→행)', N,
    sum(1 for k in range(N) if ATTR[k].get('측점') != str(rows[k]['sta'])))

# ---------------- 출력 ----------------
print('=' * 84)
print(f'{"검증 항목":56s} {"대상":>6s} {"불일치":>7s}')
print('-' * 84)
allok = True
for name, (n, bad, memo) in R.items():
    if bad: allok = False
    print(f'{"OK " if bad == 0 else "NG "}{name:53s} {n:6d} {bad:7d} {memo}')
print('=' * 84)
print('전체 판정:', 'ALL PASS' if allok else 'FAIL')
