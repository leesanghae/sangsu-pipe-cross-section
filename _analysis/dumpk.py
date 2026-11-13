import ezdxf, sys, math
d=ezdxf.readfile('/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt')
msp=d.modelspace()
ins=[e for e in msp if e.dxftype()=='INSERT' and e.dxf.name=='관로06']
ins.sort(key=lambda e:(e.dxf.insert[0],e.dxf.insert[1]))
BX=min(e.dxf.insert[0] for e in ins); BY=min(e.dxf.insert[1] for e in ins)
import collections
def key(e):
    sh=round((e.dxf.insert[0]-BX)/84.1); r=(e.dxf.insert[0]-BX)-sh*84.1
    return (sh, round(r/18.0), round((e.dxf.insert[1]-BY)/9.5))
ins.sort(key=key)
def dump(k):
    e0=ins[k]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
    at={a.dxf.tag:a.dxf.text for a in e0.attribs}
    print('##### k=%d  sta=%s  ch=%s  D=%s  exc=%s  ar=%s'%(k,at.get('측점'),at.get('지반고'),at.get('관저고'),at.get('터파기(육상)합계'),at.get('되메우기(관주위)')))
    def P(p): return (round(p[0]-ax,3), round(p[1]-ay,3))
    def rr(en):
        t=en.dxftype()
        if t=='LINE': return ('LINE',P(en.dxf.start),P(en.dxf.end))
        if t=='LWPOLYLINE': return ('LWPOLY',[(round(q[0]-ax,3),round(q[1]-ay,3)) for q in en.get_points('xy')],'closed',en.closed)
        if t=='TEXT': return ('TEXT',repr(en.dxf.text),P(en.dxf.insert),en.dxf.height,round(en.dxf.rotation,1))
        if t=='CIRCLE': return ('CIRCLE',P(en.dxf.center),round(en.dxf.radius,3))
        if t=='HATCH': return ('HATCH',en.dxf.pattern_name,en.dxf.pattern_scale,len(en.paths),[ (len(p.edges) if hasattr(p,'edges') else -1) for p in en.paths])
        if t=='INSERT': return ('INSERT',en.dxf.name,P(en.dxf.insert),en.dxf.xscale)
        return (t,)
    seen=collections.OrderedDict()
    for en in msp:
        lay=en.dxf.layer
        if not lay.startswith('CC-'): continue
        try:
            if en.dxftype() in ('TEXT','INSERT'): p=en.dxf.insert
            elif en.dxftype()=='LINE': p=en.dxf.start
            elif en.dxftype()=='CIRCLE': p=en.dxf.center
            elif en.dxftype()=='LWPOLYLINE': p=en.get_points('xy')[0]
            elif en.dxftype()=='HATCH':
                pp=None
                for pa in en.paths:
                    if getattr(pa,'vertices',None) and len(pa.vertices): pp=pa.vertices[0];break
                    if getattr(pa,'edges',None) and len(pa.edges):
                        ed=pa.edges[0]; pp=ed.center if hasattr(ed,'center') else ed.start;break
                if pp is None: continue
                p=pp
            else: continue
        except Exception: continue
        if ax-9<p[0]<ax+30 and ay-7<p[1]<ay+4:
            seen.setdefault(lay,[]).append(rr(en))
    for lay in sorted(seen):
        print('  =====',lay,len(seen[lay]))
        for v in seen[lay]: print('    ',v)
for k in [int(x) for x in sys.argv[1:]]:
    dump(k)
