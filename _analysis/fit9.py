import ezdxf, collections, sys, math
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
LINES=[e for e in IDX['CC-LEAD'] if e.dxftype()=='LINE']
def hasBH(k):
    ax,ay=ANCH[k]
    n=0
    for e in LINES:
        if not (ax-3<e.dxf.start[0]<ax+3 and ay-3<e.dxf.start[1]<ay+3): continue
        n+=1
    return n
res=[]
for k in range(655):
    t=rows[k]
    n=hasBH(k)
    H=t['gh']-t['inv']; B=t['B']
    prev=rows[k-1] if k>0 else None
    first=(k%20==0)
    cchg = prev is not None and t['code']!=prev['code']
    Bchg = prev is not None and abs((t['B'] or 0)-(prev['B'] or 0))>1e-9
    Hchg = prev is not None and abs((t['gh']-t['inv'])-(prev['gh']-prev['inv']))>1e-9
    res.append((k,first,cchg,Bchg,Hchg,n, bool(t['exc'])))
import collections as C
print('LEAD-entity count distribution:',C.Counter(r[5] for r in res))
for cond,name in [((True,True,True,True),'any'),((True,True,False,False),'first|code'),
                  ((True,True,False,True),'first|code|H'),((True,True,True,True),'x')]:
    pass
# test hypotheses for n>0
def test(pred,label):
    bad=[r[0] for r in res if (r[5]>0)!=pred(r)]
    print('%-40s mismatch=%d %s'%(label,len(bad),bad[:8]))
test(lambda r: r[1] or r[2] or r[4] or r[3], 'first|codechg|Bchg|Hchg')
test(lambda r: r[1] or r[2], 'first|codechg')
test(lambda r: r[1] or r[2] or r[4], 'first|codechg|Hchg')
test(lambda r: r[1] or r[2] or r[3], 'first|codechg|Bchg')
