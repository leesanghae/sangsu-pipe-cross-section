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
ANCH=[(e.dxf.insert[0]-5.0, e.dxf.insert[1]) for e in ins]
LINES=[e for e in msp if e.dxf.layer=='CC-LEAL' and e.dxftype()=='LINE']
DOTS =[e for e in msp if e.dxf.layer=='CC-LEAL' and e.dxftype()=='INSERT']
data=[]
for k in range(655):
    ax,ay=ANCH[k]
    vert=[e for e in LINES if abs(e.dxf.start[0]-ax-3.3)<1e-3 and abs(e.dxf.end[0]-ax-3.3)<1e-3
          and abs(e.dxf.start[1]-ay)<3 and abs(e.dxf.end[1]-ay)<3]
    if not vert: continue
    ys=[]
    for e in vert: ys += [round(e.dxf.start[1]-ay,4), round(e.dxf.end[1]-ay,4)]
    tk=[round(e.dxf.start[1]-ay,4) for e in LINES
        if abs(e.dxf.start[0]-ax-3.3)<1e-3 and abs(e.dxf.end[0]-ax-3.8)<1e-3 and abs(e.dxf.start[1]-ay)<3]
    t=rows[k]; lay=calc.PAVE.get(t['code'][2:4],[])
    tvv=sum(x[1] for x in lay if x[2]=='CUT'); zL=sum(x[1] for x in lay)
    dots=sorted([round(e.dxf.insert[1]-ay,4) for e in DOTS if abs(e.dxf.insert[0]-ax-3.3)<1e-3 and abs(e.dxf.insert[1]-ay)<3],reverse=True)
    data.append((k,t['code'],round(t['gh']-t['inv'],3),round(tvv,3),round(zL,3),max(ys),min(ys),sorted(tk,reverse=True),dots))
print('n=',len(data))
for r in data[:10]: print('k=%-4d %-5s H=%.2f tv=%.2f zL=%.2f span=%.2f..%.2f ticks=%s dots=%s'%r)
print('\n--- hypothesis: top = 2.65 - zL ---')
for nm,f in [('2.5-tv',lambda H,tv,zL:2.5-tv),('2.65-zL',lambda H,tv,zL:2.65-zL),
             ('2.35-tv+0.3',lambda H,tv,zL:2.65-tv),('2.8-tv-0.15',lambda H,tv,zL:2.65-tv),
             ('2.2+0.45-tv',lambda H,tv,zL:2.65-tv),('2.5-zL+0.15',lambda H,tv,zL:2.65-zL)]:
    bad=[r[0] for r in data if r[7] and abs(r[7][0]-round(f(r[2],r[3],r[4]),4))>1e-3]
    print(' ',nm,'mismatch',len(bad),bad[:6])
print('\n--- bottom of leader vs zL ---')
bad=[r[0] for r in data if abs(r[6]+round(r[4],4))>1e-3]
print('  bottom == -zL mismatch',len(bad),bad[:6])
print('\n--- tick spacing ---')
for r in data[:8]:
    tk=r[7]
    print('  k=%-4d ticks=%s diffs=%s'%(r[0],tk,[round(t-te,4) for t,te in zip(tk,tk[1:])]))
