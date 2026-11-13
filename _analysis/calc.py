import openpyxl, math
from collections import Counter, defaultdict
F='/home/user/uploaded_files/종평면도(상수관로) (횡단포함).xlsx'
def rnd2(x):
    return math.copysign(math.floor(abs(x)*100+0.5+1e-9)/100,x) if x else 0.0
PAVE={'01':[], '02':[('표층',.05,'CUT'),('기층',.10,'CUT'),('보조기층',.30,'SUB')],
      '04':[('보도블럭',.06,'CUT'),('모래',.04,'SAND'),('보조기층',.10,'SUB')]}
def section(gh,inv,dia,code,slope,BW,H1=0.0,H2=0.0,neg0=True):
    D=dia/1000; m=code[:2]; p=code[2:4]
    out=dict(B=0,exc=0,ar=0,ab=0,sand=0,sub=0,frost=0,asp=0,bodo=0,raw_ar=0,has=False)
    if m!='OP': return out
    H=gh-inv; Hd=H+H1; b0=D+2*BW
    lay=PAVE[p]; tv=sum(t for _,t,r in lay if r=='CUT')
    w=lambda z: b0+2*slope*(Hd-z)
    if Hd<=0: return out
    out['has']=True
    out['B']=rnd2(w(tv))
    exc=(b0+w(tv))/2*(Hd-tv)
    z=0; lays={}
    for n,t,r in lay:
        zt,zb=z,z+t; z=zb
        if r=='CUT': continue
        zt2=min(max(zt,tv),Hd); zb2=min(max(zb,tv),Hd)
        a=(w(zt2)+w(zb2))/2*(zb2-zt2) if zb2>zt2 else 0
        lays[r]=lays.get(r,0)+a
    zL=min(z,Hd) if lay else 0
    zL=max(zL,tv)
    zP=Hd-H1-D-H2
    ab=(w(zL)+w(zP))/2*(zP-zL) if zP>zL else 0
    ar=exc-sum(lays.values())-ab-math.pi*D*D/4
    out.update(exc=rnd2(exc),ab=rnd2(ab),sand=rnd2(lays.get('SAND',0)),sub=rnd2(lays.get('SUB',0)),frost=rnd2(lays.get('FROST',0)),raw_ar=ar)
    out['ar']=rnd2(ar) if (ar>=0 or not neg0) else 0.0
    if p in('02',): out['asp']=out['B']
    if p=='04': out['bodo']=out['B']
    return out
def load():
    wb=openpyxl.load_workbook(F,data_only=True); T=wb['토적표']; rows=[]
    for r in range(6,661):
        g=lambda c: T.cell(r,c).value
        rows.append(dict(r=r,sta=g(2),dist=g(3),seg=g(4),gh=g(5),inv=g(7),eq=g(8),exc=g(9),wet=g(10),ar=g(11),ab=g(12),asp=g(13),con=g(14),asc=g(15),bodo=g(16),sand=g(17),sub=g(18),frost=g(19),crush=g(20),
          dia=g(96),dep=g(98),B=g(99),code=g(100),V=[g(c) for c in range(21,91)]))
    return rows,wb
if __name__=='__main__':
    rows,wb=load()
    for BW160 in (0.0,0.2):
        bad=defaultdict(list); n=0
        for t in rows:
            BW=BW160 if t['dia']==160 else 0.2
            o=section(t['gh'],t['inv'],t['dia'],t['code'],0.3,BW)
            for k in ('B','exc','ar','ab','sand','sub','frost','asp','bodo'):
                if abs(o[k]-(t[k] or 0))>0.0051: bad[k].append((t['r'],t['sta'],t['dia'],t['code'],t[k],o[k],round(o['raw_ar'],4) if k=='ar' else ''))
        print(f'=== BW(160)={BW160}: mismatching rows per item', {k:len(v) for k,v in bad.items()})
        for k,v in bad.items():
            print('  ',k,'diff distribution',Counter(round(b[5]-b[4],2) for b in v).most_common(6)); 
            for b in v[:6]: print('     ',b)
