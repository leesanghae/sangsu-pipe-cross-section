import ezdxf, collections, sys
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
PAVE=[e for e in msp if e.dxf.layer=='CC-PAVE' and e.dxftype()=='LWPOLYLINE']
bad=[]
for k in range(655):
    t=rows[k]; ax,ay=ANCH[k]
    got=set()
    for e in PAVE:
        p=[(q[0],q[1]) for q in e.get_points('xy')]
        if abs(p[0][0]-ax)>1e-6: continue
        if not (ax-5<p[0][0]<ax+5 and ay-2<p[0][1]<ay+1): continue
        if abs(abs(p[1][0]-ax)-4.0)>1e-3: continue
        got.add(round(p[0][1]-ay,4))
    lay=calc.PAVE.get(t['code'][2:4],[])
    want=set(); z=0.0
    for n,th,kd in lay:
        z+=th; want.add(-round(z,4))
    if got!=want: bad.append((k,t['code'],sorted(got),sorted(want)))
print('pavement SET mismatch:',len(bad))
for b in bad[:8]: print('  ',b)
