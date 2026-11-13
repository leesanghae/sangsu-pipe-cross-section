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

obs=[]
for k in range(655):
    t=rows[k]; ax,ay=ANCH[k]
    left=collections.Counter(); right=collections.Counter()
    for e in PAVE:
        p=[(q[0],q[1]) for q in e.get_points('xy')]
        if not(ax-5<p[0][0]<ax+5 and ay-2<p[0][1]<ay+1): continue
        y=round(p[0][1]-ay,4)
        side = 'L' if p[1][0]<p[0][0] else 'R'
        if abs(p[0][0]-ax)>1e-6: continue
        (left if side=='L' else right)[y]+=1
    lay=calc.PAVE.get(t['code'][2:4],[])
    cum=[]; z=0.0
    for n,th,kd in lay:
        z+=th; cum.append((round(z,4),kd,n))
    obs.append((k,t['code'],cum,dict(left),dict(right)))
print('  k code  layers(cum,kind)                          LEFT           RIGHT')
for k,code,cum,L,R in obs[:6]+[o for o in obs if o[1][:2]=='OP04'][:3]:
    print('%4d %-5s %-40s L=%-22s R=%s'%(k,code,cum,L,R))

# hypothesis A: each layer bottom once per side
def predA(cum):
    C=collections.Counter()
    for z,kd,n in cum: C[-z]+=1
    return C
# hypothesis B: bottom of each layer + repeat for CUT layers
def predB(cum):
    C=collections.Counter()
    for z,kd,n in cum:
        C[-z]+=1
        if kd=='CUT': C[-z]+=1
    return C
# hypothesis C: bottom + tv + zL
def predC(cum):
    C=collections.Counter()
    for z,kd,n in cum: C[-z]+=1
    tv=sum(z for z,kd,n in cum if kd=='CUT')
    if cum: 
        C[-round(tv,4)]+=1; C[-cum[-1][0]]+=1
    return C
# hypothesis D: bottom + top-of-each-layer(except ground)
def predD(cum):
    C=collections.Counter(); prev=0.0
    for z,kd,n in cum:
        C[-z]+=1
        if abs(prev)>1e-9: C[-round(prev,4)]+=1
        prev=z
    return C
for nm,f in [('A cumulative only',predA),('B cum+CUTdup',predB),('C cum+tv+zL',predC),('D cum+tops',predD)]:
    bad=0; badk=[]
    for k,code,cum,L,R in obs:
        p=f(cum)
        if p!=collections.Counter(L) or p!=collections.Counter(R):
            bad+=1
            if len(badk)<5: badk.append(k)
    print('%-20s mismatch=%3d %s'%(nm,bad,badk))
