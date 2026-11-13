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
IDX=collections.defaultdict(list)
for e in msp: IDX[e.dxf.layer].append(e)
ANCH=[(e.dxf.insert[0]-5.0, e.dxf.insert[1]) for e in ins]
def cnt(k,lay):
    ax,ay=ANCH[k]; n=0
    for e in IDX[lay]:
        t=e.dxftype()
        try:
            if t in ('TEXT','INSERT'): p=e.dxf.insert
            elif t=='LINE': p=e.dxf.start
            elif t=='CIRCLE': p=e.dxf.center
            elif t=='LWPOLYLINE': p=e.get_points('xy')[0]
            elif t=='HATCH':
                pp=None
                for pa in e.paths:
                    if getattr(pa,'vertices',None) and len(pa.vertices): pp=pa.vertices[0];break
                    if getattr(pa,'edges',None) and len(pa.edges):
                        ed=pa.edges[0]; pp=ed.center if hasattr(ed,'center') else ed.start;break
                if pp is None: continue
                p=pp
            else: continue
        except Exception: continue
        if ax-9<p[0]<ax+10 and ay-7<p[1]<ay+4: n+=1
    return n
print('%-4s %-14s %-5s %5s %4s %5s %5s | %s'%('k','sta','code','H','exc','ar','B','CENT GRND PAVE EXCV PIPE RFIL PATT LEAD LEAL TABL'))
for k in [104,160,163,7,8,9]:
    t=rows[k]; H=t['gh']-t['inv']
    print('%-4d %-14s %-5s %5.2f %4s %5s %5s | %s'%(k,t['sta'],t['code'],H,
      ('Y' if t['exc'] else '.'),t['ar'],t['B'],
      ' '.join('%d'%cnt(k,L) for L in ('CC-CENT','CC-GRND','CC-PAVE','CC-EXCV','CC-PIPE','CC-RFIL','CC-PATT','CC-LEAD','CC-LEAL','CC-TABL'))))
