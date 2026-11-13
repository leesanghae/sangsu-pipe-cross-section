import ezdxf, math, collections
d=ezdxf.readfile('/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt')
msp=d.modelspace()
ins=[e for e in msp if e.dxftype()=='INSERT' and e.dxf.name=='관로06']
BX=min(e.dxf.insert[0] for e in ins); BY=min(e.dxf.insert[1] for e in ins)
def key(e):
    sh=round((e.dxf.insert[0]-BX)/84.1); r=(e.dxf.insert[0]-BX)-sh*84.1
    return (sh, round(r/18.0), round((e.dxf.insert[1]-BY)/9.5))
ins.sort(key=key)
HALIGN={0:'L',1:'C',2:'R',3:'A',4:'M',5:'F'}
VALIGN={0:'B',1:'B',2:'M',3:'T'}
def dump(k, win=(-9,10,-7,4)):
    e0=ins[k]; ax=e0.dxf.insert[0]-5.0; ay=e0.dxf.insert[1]
    at={a.dxf.tag:a.dxf.text for a in e0.attribs}
    print('#### k=%d sta=%s GH=%s INV=%s D=%s exc=%s'%(k,at.get('측점'),at.get('지반고'),at.get('관저고'),at.get('터파기(육상)합계'),'?'))
    for en in msp:
        if en.dxf.layer not in ('CC-TEXT','CC-LEAT','CC-DIML'): continue
        if en.dxftype()!='TEXT': continue
        p=en.dxf.insert
        if not (ax+win[0]<p[0]<ax+win[1] and ay+win[2]<p[1]<ay+win[3]): continue
        print('   %-8s %-22r rel=(%.3f, %.3f) h=%.3f rot=%.1f ha=%s va=%s'%(
            en.dxf.layer, en.dxf.text, p[0]-ax, p[1]-ay, en.dxf.height, en.dxf.rotation,
            HALIGN.get(en.dxf.halign,'?'), VALIGN.get(en.dxf.valign,'?')))
for k in [0,20,40]:
    dump(k)
