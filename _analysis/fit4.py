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
def tv(code): return sum(t for _,t,r in calc.PAVE.get(code[2:4],[]) if r=='CUT')
def bw(dia): return 0.0 if dia==160 else 0.20

# ---- HATCH edges for k=0 ----
e0=ins[0]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
for e in IDX['CC-PATT']:
    p=e.paths[0]; 
    try:
        pp=p.edges[0].start if hasattr(p.edges[0],'start') else p.edges[0].center
    except Exception: pp=None
    if pp is None or not(ax-9<pp[0]<ax+9 and ay-7<pp[1]<ay+4): continue
    print('HATCH k=0 paths=%d solid=%s assoc=%s'%(len(e.paths),e.dxf.solid_fill,e.dxf.associative))
    for i,pa in enumerate(e.paths):
        print('  path',i,'edges',len(pa.edges))
        for ed in pa.edges:
            t=ed.EDGE_TYPE
            if t=='LineEdge':
                print('    LINE (%.4f,%.4f)->(%.4f,%.4f)'%(ed.start[0]-ax,ed.start[1]-ay,ed.end[0]-ax,ed.end[1]-ay))
            else:
                print('    ARC c=(%.4f,%.4f) r=%.4f a=%.1f..%.1f ccw=%s'%(ed.center[0]-ax,ed.center[1]-ay,ed.radius,ed.start_angle,ed.end_angle,ed.ccw))
    break

# ---- LEAL tick anchor over all sections that have LEAL ----
print('\n---- LEAL ----')
cnt=0
for k in range(655):
    t=rows[k]
    e0=ins[k]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
    hits=[e for e in IDX['CC-LEAL'] if e.dxftype()=='LINE']
    loc=[]
    for e in hits:
        if ax+3.2<e.dxf.start[0]<ax+3.4:
            loc.append(round(e.dxf.start[1]-ay,4))
    if not loc: continue
    H=t['gh']-t['inv']; D=t['dia']/1000.0
    lay=calc.PAVE.get(t['code'][2:4],[])
    tvv=tv(t['code'])
    mids=[]; z=0.0
    for n,th,ty in lay:
        mids.append(round(-(z+th/2),4)); z+=th
    loc.sort(reverse=True)
    if cnt<20: print('k=%-4d %-14s %-5s tv=%.2f H=%.2f  ticks=%s  mids=%s'%(k,t['sta'],t['code'],tvv,H,loc,mids))
    cnt+=1
print('total LEAL sections',cnt)
