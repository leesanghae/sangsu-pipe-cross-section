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
LEAD=[e for e in msp if e.dxf.layer=='CC-LEAD']
def dump(k):
    ax,ay=ANCH[k]; t=rows[k]
    H=t['gh']-t['inv']
    print('#### k=%d %s code=%s H=%.2f B=%s exc=%s'%(k,t['sta'],t['code'],H,t['B'],t['exc']))
    out=[]
    for e in LEAD:
        if e.dxftype()=='LINE':
            if ax-3<e.dxf.start[0]<ax+3 and ay-3<e.dxf.start[1]<ay+3:
                out.append(('L',round(e.dxf.start[0]-ax,3),round(e.dxf.start[1]-ay,3),round(e.dxf.end[0]-ax,3),round(e.dxf.end[1]-ay,3)))
        elif e.dxftype()=='INSERT':
            if ax-3<e.dxf.insert[0]<ax+3 and ay-3<e.dxf.insert[1]<ay+3:
                out.append(('D',round(e.dxf.insert[0]-ax,3),round(e.dxf.insert[1]-ay,3)))
    for o in sorted(out,key=lambda r:(-r[2],r[1])): print('   ',o)
for k in [1,163,2]: dump(k)
print()
print('### k=356 polygon / pipe')
k=356; t=rows[k]; e0=ins[k]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
print('tokjuk dia',t['dia'],'H',t['gh']-t['inv'],'B',t['B'],'exc',t['exc'],'ar',t['ar'])
for e in msp:
    if e.dxf.layer=='CC-EXCV' and e.dxftype()=='LWPOLYLINE':
        p=[(q[0]-ax,q[1]-ay) for q in e.get_points('xy')]
        if ax-9<p[0][0]<ax+10 and ay-7<p[0][1]<ay+4: print('  POLY',[(round(a,3),round(b,3)) for a,b in p])
    if e.dxf.layer=='CC-PIPE' and e.dxftype()=='CIRCLE':
        if ax-9<e.dxf.center[0]<ax+10 and ay-7<e.dxf.center[1]<ay+4: print('  CIRC r=%.3f c=(%.3f,%.3f)'%(e.dxf.radius,e.dxf.center[0]-ax,e.dxf.center[1]-ay))
