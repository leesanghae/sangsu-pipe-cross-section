import ezdxf, math
from collections import Counter
doc=ezdxf.readfile('/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt')
msp=doc.modelspace()
X0=2469.908; Y0=4655.029
def cell(k):
    sh=k//20; col=(k%20)//5; row=k%5
    return X0+sh*84.1+13.0+col*18.0, Y0+11.5+row*9.5
ENTS=[]
for e in msp:
    t=e.dxftype(); lay=e.dxf.layer
    if not lay.startswith('CC-') or lay in('CC-BORDER','CC-Title'): continue
    try:
        if t=='LINE': x,y=e.dxf.start[0],e.dxf.start[1]
        elif t=='LWPOLYLINE': x,y=list(e.get_points('xy'))[0]
        elif t in('TEXT','INSERT'): x,y=e.dxf.insert[0],e.dxf.insert[1]
        elif t=='CIRCLE': x,y=e.dxf.center[0],e.dxf.center[1]
        elif t=='HATCH':
            x=y=None
            for p in e.paths:
                if hasattr(p,'vertices') and p.vertices: x,y=p.vertices[0][0],p.vertices[0][1]; break
                if hasattr(p,'edges') and p.edges:
                    ed=p.edges[0]
                    if hasattr(ed,'start'): x,y=ed.start[0],ed.start[1]
                    elif hasattr(ed,'center'): x,y=ed.center[0],ed.center[1]
                    break
            if x is None: continue
        else: continue
    except Exception: continue
    ENTS.append((x,y,t,lay,e))
def near(k):
    cx,gy=cell(k)
    return cx,gy,[E for E in ENTS if cx-8<=E[0]<=cx+10 and gy-5.6<=E[1]<=gy+3.9]
def desc(E,cx,gy):
    x,y,t,lay,e=E
    f=lambda a,b:f"({a-cx:.3f},{b-gy:.3f})"
    if t=='LINE': return f"LINE {f(e.dxf.start[0],e.dxf.start[1])}->{f(e.dxf.end[0],e.dxf.end[1])}"
    if t=='LWPOLYLINE': return "LWPOLY "+("closed " if e.closed else "")+" ".join(f(a,b) for a,b in e.get_points('xy'))
    if t=='TEXT':
        ap=e.dxf.get('align_point')
        return f"TEXT {e.dxf.text!r} h={e.dxf.height} rot={round(e.dxf.get('rotation',0),1)} ha={e.dxf.get('halign',0)} va={e.dxf.get('valign',0)} at="+(f(ap[0],ap[1]) if ap is not None else f(e.dxf.insert[0],e.dxf.insert[1]))
    if t=='CIRCLE': return f"CIRCLE c={f(e.dxf.center[0],e.dxf.center[1])} r={e.dxf.radius:.4f}"
    if t=='INSERT': return f"INSERT {e.dxf.name} {f(e.dxf.insert[0],e.dxf.insert[1])} sc={e.dxf.xscale}"
    return t
def dump(k,label,skip=('CC-PATT','CC-TABL')):
    cx,gy,es=near(k)
    print('\n==== %s  k=%d'%(label,k))
    for d in sorted(f"[{E[3]}] {desc(E,cx,gy)}" for E in es if E[3] not in skip): print('  '+d)
    print('  counts:',dict(Counter(E[3] for E in es)))
