# -*- coding: utf-8 -*-
import ezdxf
d=ezdxf.readfile('/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt')

def blk(nm):
    for b in d.blocks:
        if b.name==nm: return b
    return None

def fmt(x):
    s=('%.6f'%x).rstrip('0').rstrip('.')
    return s if s not in ('','-0') else '0'

def emit_table():
    b=blk('관로06')
    L=['Public Sub CreateTableBlock()',
       '    \' 관로06 : 70 x 35.2 표 블록 (원본 DXF 정의 그대로)',
       '    Dim blk As Object, p3 As Variant, e As Object, att As Object',
       '    If BlockExists(pBlkTable) Then Exit Sub',
       '    Set blk = doc.Blocks.Add(P3(0#, 0#), pBlkTable)',
       '    \' 외곽 (closed LWPOLYLINE)']
    poly=None; lines=[]; texts=[]; atts=[]
    for e in b:
        t=e.dxftype()
        if t=='LWPOLYLINE': poly=[(q[0],q[1]) for q in e.get_points('xy')]
        elif t=='LINE': lines.append((e.dxf.start[0],e.dxf.start[1],e.dxf.end[0],e.dxf.end[1]))
        elif t=='TEXT': texts.append((e.dxf.text,e.dxf.insert[0],e.dxf.insert[1],e.dxf.height))
        elif t=='ATTDEF':
            ap=e.dxf.align_point
            ax,ay=(ap[0],ap[1]) if ap else (e.dxf.insert[0],e.dxf.insert[1])
            atts.append((e.dxf.tag,ax,ay,e.dxf.height,e.dxf.halign,e.dxf.valign))
    if poly:
        a=[]
        for (x,y) in poly: a += [fmt(x),fmt(y)]
        L.append('    Dim a0(0 To %d) As Double'%(len(a)-1))
        for i,v in enumerate(a):
            L.append('    a0(%d) = %s'%(i,v))
        L.append('    Set e = blk.AddLightWeightPolyline(a0)')
        L.append('    e.Closed = True')
    L.append('    \' 내부 구분선 %d개'%len(lines))
    for (x1,y1,x2,y2) in lines:
        L.append('    blk.AddLine P3(%s, %s), P3(%s, %s)'%(fmt(x1),fmt(y1),fmt(x2),fmt(y2)))
    L.append('    \' 고정 문자 %d개'%len(texts))
    for (s,x,y,h) in texts:
        L.append('    blk.AddText "%s", P3(%s, %s), %s'%(s.replace('"','""'),fmt(x),fmt(y),fmt(h)))
    L.append('    \' 속성 정의 %d개'%len(atts))
    for (tag,x,y,h,ha,va) in atts:
        L.append('    Set att = blk.AddAttribute(%s, 0, "%s", P3(%s, %s), "%s", "")'%(fmt(h),tag,fmt(x),fmt(y),tag))
        al = align_of(ha,va)
        if al:
            L.append('    att.Alignment = %s'%al)
            L.append('    att.TextAlignmentPoint = P3(%s, %s)'%(fmt(x),fmt(y)))
    L.append('End Sub')
    return L

def align_of(ha,va):
    # DXF halign/valign -> acAlignment
    m={(0,0):None,(0,2):None,(1,2):'acAlignmentMiddleCenter',(1,0):'acAlignmentBottomCenter',
       (2,2):'acAlignmentMiddleRight',(0,2):None,(0,1):None}
    if (ha,va)==(1,2): return 'acAlignmentMiddleCenter'
    if (ha,va)==(1,0): return 'acAlignmentBottomCenter'
    if (ha,va)==(2,2): return 'acAlignmentMiddleRight'
    if (ha,va)==(0,2): return None
    if (ha,va)==(0,0): return None
    if (ha,va)==(1,1): return 'acAlignmentMiddleCenter'
    return None

print('\n'.join(emit_table()[:20]))
print('...')
