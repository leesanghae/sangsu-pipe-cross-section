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
for e in msp:
    IDX[e.dxf.layer].append(e)
def pos(e):
    ap=e.dxf.align_point
    return (ap[0],ap[1]) if ap else (e.dxf.insert[0],e.dxf.insert[1])
def tv(code): return sum(t for _,t,r in calc.PAVE.get(code[2:4],[]) if r=='CUT')
def bw(dia): return 0.0 if dia==160 else 0.20

for k in [0,100,180]:
    t=rows[k]; e0=ins[k]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
    H=t['gh']-t['inv']; D=t['dia']/1000.0; b0=D+2*bw(t['dia']); tvv=tv(t['code'])
    zP=H-D; 
    print('#### k=%d sta=%s D=%d H=%.2f b0=%.2f tv=%.2f zP=%.2f'%(k,t['sta'],t['dia'],H,b0,tvv,zP))
    print('   LEAD:')
    for e in IDX['CC-LEAD']:
        p=pos(e)
        if ax-9<p[0]<ax+9 and ay-7<p[1]<ay+4.2:
            if e.dxftype()=='LINE':
                print('     LINE (%.4f,%.4f)->(%.4f,%.4f)'%(e.dxf.start[0]-ax,e.dxf.start[1]-ay,e.dxf.end[0]-ax,e.dxf.end[1]-ay))
            else: print('     %s (%s)'%(e.dxftype(),'HDot'))
    print('   LEAT texts:')
    for e in IDX['CC-LEAT']:
        if e.dxftype()!='TEXT': continue
        p=pos(e)
        if ax-9<p[0]<ax+9 and ay-7<p[1]<ay+4.2:
            print('     %-18r align=(%.4f,%.4f) h=%.2f rot=%.0f ha=%d va=%d'%(e.dxf.text,p[0]-ax,p[1]-ay,e.dxf.height,e.dxf.rotation,e.dxf.halign,e.dxf.valign))
