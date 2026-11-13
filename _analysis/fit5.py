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
def bw(dia): return 0.0 if dia==160 else 0.20

# --- CC-Title inserts ---
print('=== CC-Title / CC-BORDER inserts ===')
for e in IDX['CC-Title'][:3]:
    print(' TITLE ins=(%.3f,%.3f) sc=(%.3f,%.3f) rot=%.1f layer=%s'%(e.dxf.insert[0],e.dxf.insert[1],e.dxf.xscale,e.dxf.yscale,e.dxf.rotation,e.dxf.layer))
    print('    attrs:',{a.dxf.tag:a.dxf.text for a in e.attribs})
for e in IDX['CC-BORDER'][:2]:
    try: print(' BORDER %s ins=(%.4f,%.4f) sc=%.4f'%(e.dxftype(),e.dxf.insert[0],e.dxf.insert[1],e.dxf.xscale))
    except Exception as ex: print(' BORDER',e.dxftype(),ex)

# --- 도면번호 block ATTDEF positions/tags ---
for b in d.blocks:
    if b.name=='도면번호':
        print('=== 도면번호 block base=',b.block.dxf.base_point)
        for e in b:
            if e.dxftype()=='ATTDEF':
                print('   ATTDEF %-10s p=(%.3f,%.3f) h=%.2f ha=%d va=%d style=%s'%(e.dxf.tag,e.dxf.insert[0],e.dxf.insert[1],e.dxf.height,e.dxf.halign,e.dxf.valign,e.dxf.style))
            elif e.dxftype() in ('LINE','LWPOLYLINE'):
                print('   ',e.dxftype())

# --- bottom-dim texts for D160 first-sections ---
print('=== bottom dim x for first sections ===')
for k in range(0,655,20):
    t=rows[k]
    e0=ins[k]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
    H=t['gh']-t['inv']; D=t['dia']/1000.0; b0=D+2*bw(t['dia'])
    got=[]
    for e in IDX['CC-LEAT']:
        if e.dxftype()!='TEXT': continue
        ap=e.dxf.align_point
        if ap is None: continue
        y=ap[1]-ay; x=ap[0]-ax
        if abs(y+(H+1.6))<0.02 or abs(y+(H+1.2))<0.02:
            got.append((round(x,4),round(y,4),e.dxf.text))
    got.sort()
    print('k=%-3d D=%3d b0=%.2f  %s'%(k,t['dia'],b0,got))
