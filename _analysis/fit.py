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
HALIGN={0:'L',1:'C',2:'R',3:'A',4:'M',5:'F'}
# collect TEXT on CC-LEAT/CC-TEXT/CC-DIML near each anchor
IDX=collections.defaultdict(list)
for e in msp:
    if e.dxftype()!='TEXT': continue
    IDX[e.dxf.layer].append(e)
def tv(code): return sum(t for _,t,r in calc.PAVE.get(code[2:4],[]) if r=='CUT')
recs=[]
for k in range(0,655,1):
    t=rows[k]
    if k%20: continue
    e0=ins[k]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
    txts=[]
    for lay in ('CC-LEAT','CC-TEXT','CC-DIML'):
        for e in IDX[lay]:
            p=e.dxf.insert
            if ax-9<p[0]<ax+9 and ay-7<p[1]<ay+4.2:
                txts.append((lay,e.dxf.text,round(p[0]-ax,4),round(p[1]-ay,4),e.dxf.height,round(e.dxf.rotation,1),HALIGN.get(e.dxf.halign)))
    H=t['gh']-t['inv']; D=t['dia']/1000.0
    recs.append((k,t['sta'],t['code'],H,D,tv(t['code']),t['B'],txts))
print('n=',len(recs))
# print the B/H texts and the outer ones
for k,sta,code,H,D,tvv,B,txts in recs:
    L=[x for x in txts if x[1].startswith(('B=','H='))]
    O=[x for x in txts if x[0]=='CC-LEAT' and x[4]==0.2 and x[5]==90]
    Bo=[x for x in txts if x[0]=='CC-LEAT' and x[4]==0.2 and x[5]==0 and not x[1].startswith(('표','기','보','모','되'))]
    print('k=%-3d %-14s %-5s H=%.2f D=%.2f tv=%.2f B=%s' % (k,sta,code,H,D,tvv,B))
    for x in L: print('     %-14r rel=(%.4f,%.4f) rot=%g' % (x[1],x[2],x[3],x[5]))
    for x in O: print('     OUTER  %-8r rel=(%.4f,%.4f)' % (x[1],x[2],x[3]))
    for x in Bo: print('     BOT    %-8r rel=(%.4f,%.4f)' % (x[1],x[2],x[3]))
