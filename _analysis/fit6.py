import ezdxf, math, collections, sys
sys.path.insert(0,'.')
import calc
rows,_=calc.load()
d=ezdxf.readfile('/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt')
msp=d.modelspace()
ins=[e for e in msp if e.dxftype()=='INSERT' and e.dxf.name=='관로06']
BX=min(e.dxf.insert[0] for e in ins); BY=min(e.dxf.insert[1] for e in ins)
def key(e):
    sh=round((e.dxf.insert[0]-BX)/84.1); r=(e.dxf.insert[0]-BX)-sh*84.1
    return (sh, round(r/18.0), round((e.dxf.insert[1]-BY)/9.5))
ins.sort(key=key)
IDX=collections.defaultdict(list)
for e in msp: IDX[e.dxf.layer].append(e)
ANCH=[(e.dxf.insert[0]-5.0, e.dxf.insert[1]) for e in ins]
def near(lay,k,x0,x1,y0,y1):
    ax,ay=ANCH[k]
    out=[]
    for e in IDX[lay]:
        if e.dxftype()=='LINE': p=(e.dxf.start[0],e.dxf.start[1])
        elif e.dxftype()=='INSERT': p=(e.dxf.insert[0],e.dxf.insert[1])
        elif e.dxftype()=='TEXT':
            ap=e.dxf.align_point; p=(ap[0],ap[1]) if ap else (e.dxf.insert[0],e.dxf.insert[1])
        elif e.dxftype()=='LWPOLYLINE': p=e.get_points('xy')[0]
        else: continue
        if ax+x0<=p[0]<=ax+x1 and ay+y0<=p[1]<=ay+y1: out.append((p[0]-ax,p[1]-ay,e))
    return out

print('=== LEAL: vertical line ends + tick tops (only sections where x~+3.3 belongs to this cell) ===')
data=[]
for k in range(655):
    ax,ay=ANCH[k]
    v=[e for (x,y,e) in IDX['CC-LEAL'] if e.dxftype()=='LINE'
       and abs(e.dxf.start[0]-ax-3.3)<1e-3 and abs(e.dxf.end[0]-ax-3.3)<1e-3]
    if not v: continue
    ys=[round(e.dxf.start[1]-ay,4) for e in v]+[round(e.dxf.end[1]-ay,4) for e in v]
    tk=[round(e.dxf.start[1]-ay,4) for e in IDX['CC-LEAL'] if e.dxftype()=='LINE'
        and abs(e.dxf.start[0]-ax-3.3)<1e-3 and abs(e.dxf.end[0]-ax-3.8)<1e-3]
    t=rows[k]; H=t['gh']-t['inv']
    lay=calc.PAVE.get(t['code'][2:4],[])
    tvv=sum(x[1] for x in lay if x[2]=='CUT')
    zL=sum(x[1] for x in lay)
    data.append((k,t['code'],H,tvv,zL,max(ys),min(ys),sorted(tk,reverse=True)))
for r in data[:12]: print('k=%-4d %-5s H=%.2f tv=%.2f zL=%.2f  span=%.2f..%.2f  ticks=%s'%(r[0],r[1],r[2],r[3],r[4],r[7][0] if r[7] else 9,r[6] if False else r[5],r[7]))
print('...')
# fit: top of leader = ? try 2.5 - tv
for hyp in ('2.5-tv','2.2+tv','2.35-0.5*tv','2.4-H*0+0','2.2+0.15-tv'):
    pass
print('\n=== try top = 2.5 - tv ===')
bad=0
for k,code,H,tvv,zL,mx,mn,tk in data:
    pred=round(2.5-tvv,4)
    if tk and abs(tk[0]-pred)>1e-3: bad+=1
print('mismatch',bad,'of',len(data))
print('=== try top = 2.65 - zL ===')
bad=0
for k,code,H,tvv,zL,mx,mn,tk in data:
    if tk and abs(tk[0]-round(2.65-zL,4))>1e-3: bad+=1
print('mismatch',bad,'of',len(data))
print('=== try top = 2.8 - H/2 ===')
bad=0
for k,code,H,tvv,zL,mx,mn,tk in data:
    if tk and abs(tk[0]-round(2.8-H/2,4))>1e-3: bad+=1
print('mismatch',bad,'of',len(data))
print('=== report (top, tv, zL, H) samples ===')
for k,code,H,tvv,zL,mx,mn,tk in data[:14]:
    print('  k=%-4d %-5s H=%.2f tv=%.2f zL=%.2f top=%.2f  top+tv=%.2f top+zL=%.2f'%(k,code,H,tvv,zL,tk[0] if tk else -9,(tk[0]+tvv) if tk else -9,(tk[0]+zL) if tk else -9))

print('\n=== CC-TEXT box polyline ===')
for k in [0,20,40]:
    for (x,y,e) in near('CC-TEXT',k,-2,2,-6,-4.2):
        if e.dxftype()=='LWPOLYLINE':
            print('k=%-3d box %s closed=%s'%(k,[(round(q[0]-ANCH[k][0],4),round(q[1]-ANCH[k][1],4)) for q in e.get_points('xy')],e.closed))
