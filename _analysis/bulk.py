import sys, pickle, math; sys.path.insert(0,'.')
from dx import *
from calc import load, section, rnd2, PAVE
from collections import Counter, defaultdict
tabs=pickle.load(open('tabs.pkl','rb'))
rows,wb=load()
assert len(tabs)==len(rows)==655
# bucket entities per cell for speed
B=defaultdict(list)
for E in ENTS:
    x,y=E[0],E[1]
    sh=int((x-X0)//84.1); lx=x-X0-sh*84.1; ly=y-Y0
    col=int((lx-13.0+8.0)//18.0); row=int((ly-11.5+5.6)//9.5)
    if 0<=col<4 and 0<=row<5 and sh>=0: B[sh*20+col*5+row].append(E)
res=[]
for k in range(655):
    cx,gy=cell(k); t=rows[k]; a=tabs[k][3]; es=B[k]
    d=dict(k=k,sta=t['sta'],code=t['code'],dia=t['dia'],gh=t['gh'],inv=t['inv'])
    d['att_ok']=(a['측점']==str(t['sta']) and abs(float(a['지반고'])-t['gh'])<0.006 and abs(float(a['관저고'])-t['inv'])<0.006)
    d['cover']=a['토피']
    lay=Counter(E[3] for E in es); d['lay']=lay
    d['hatch']=sum(1 for E in es if E[2]=='HATCH')
    el=sorted([(round(E[4].dxf.get('align_point')[1]-gy,3),E[4].dxf.text,round(E[4].dxf.get('align_point')[0]-cx,3)) for E in es if E[3]=='CC-TEXT' and E[2]=='TEXT' and E[4].dxf.text.lstrip('-').isdigit()])
    d['el']=el
    d['label']=[E[4].dxf.text for E in es if E[3]=='CC-TEXT' and E[2]=='TEXT' and not E[4].dxf.text.lstrip('-').isdigit()]
    d['ticks']=sum(1 for E in es if E[3]=='CC-DIML' and E[2]=='LINE' and abs(E[4].dxf.start[1]-E[4].dxf.end[1])<1e-6)
    rf=[E[4] for E in es if E[3]=='CC-RFIL']
    d['rfil']=[(round(e.dxf.start[0]-cx,3),round(e.dxf.start[1]-gy,3),round(e.dxf.end[0]-cx,3)) for e in rf]
    ci=[E[4] for E in es if E[3]=='CC-PIPE']
    d['pipe']=[(round(e.dxf.center[1]-gy,3),round(e.dxf.radius,3)) for e in ci]
    ex=[E[4] for E in es if E[3]=='CC-EXCV']
    d['excv']=[[(round(p[0]-cx,3),round(p[1]-gy,3)) for p in e.get_points('xy')] for e in ex]
    d['txt']=[E[4].dxf.text for E in es if E[3]=='CC-LEAT']
    d['leal']=lay.get('CC-LEAL',0); d['lead']=lay.get('CC-LEAD',0)
    d['grnd']=[[(round(p[0]-cx,3),round(p[1]-gy,3)) for p in E[4].get_points('xy')] for E in es if E[3]=='CC-GRND']
    d['pave']=sorted(set(round(list(E[4].get_points('xy'))[0][1]-gy,3) for E in es if E[3]=='CC-PAVE'))
    res.append(d)
pickle.dump(res,open('bulk.pkl','wb'))
print('saved',len(res),'| unassigned check: total ents',len(ENTS),'assigned',sum(len(v) for v in B.values()))
