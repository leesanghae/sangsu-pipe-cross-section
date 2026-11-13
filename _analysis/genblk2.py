# -*- coding: utf-8 -*-
import ezdxf, sys
d=ezdxf.readfile('/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt')

def fmt(x):
    s=('%.6f'%float(x)).rstrip('0').rstrip('.')
    if s in ('','-0'): s='0'
    return s

def A(ha,va):
    return {(0,0):0,(0,1):12,(0,2):9,(0,3):6,
            (1,0):13,(1,1):13,(1,2):10,(1,3):7,
            (2,0):14,(2,1):14,(2,2):11,(2,3):8}.get((ha,va),0)

def attr_of(e):
    ap=e.dxf.align_point
    return (ap[0],ap[1]) if ap else (e.dxf.insert[0],e.dxf.insert[1])

def blk(nm):
    for b in d.blocks:
        if b.name==nm: return b

out=[]
W=out.append

# ---------------- 관로06 ----------------
b=blk('관로06')
poly=[(q[0],q[1]) for e in b if e.dxftype()=='LWPOLYLINE' for q in e.get_points('xy')]
lines=[(e.dxf.start[0],e.dxf.start[1],e.dxf.end[0],e.dxf.end[1]) for e in b if e.dxftype()=='LINE']
texts=[(e.dxf.text,e.dxf.insert[0],e.dxf.insert[1],e.dxf.height,e.dxf.halign,e.dxf.valign,attr_of(e))
       for e in b if e.dxftype()=='TEXT']
atts=[(e.dxf.tag,attr_of(e)[0],attr_of(e)[1],e.dxf.height,e.dxf.halign,e.dxf.valign)
      for e in b if e.dxftype()=='ATTDEF']

W("'=====================================================================================")
W("'  블록 자체 생성 (원본 DWG 에 블록이 없어도 동작하도록)")
W("'  ------------------------------------------------------------------------------")
W("'  좌표는 원본 DXF 블록정의에서 그대로 추출한 값이다. 임의로 고치지 말 것.")
W("'=====================================================================================")
W("Private Function BlockExists(ByVal nm As String) As Boolean")
W("    Dim b As Object")
W("    BlockExists = False")
W("    If Len(nm) = 0 Then Exit Function")
W("    On Error Resume Next")
W("    Set b = doc.Blocks.Item(nm)")
W("    If Err.Number <> 0 Then Err.Clear: Exit Function")
W("    On Error GoTo 0")
W("    BlockExists = Not (b Is Nothing)")
W("End Function")
W("")
W("' 관로06 : 70.0 x 35.2 표 블록  (LINE %d, TEXT %d, ATTDEF %d)"%(len(lines),len(texts),len(atts)))
W("Public Sub CreateTableBlock()")
W("    Dim blk As Object, e As Object, at As Object")
W("    If BlockExists(pBlkTable) Then Exit Sub")
W("    On Error Resume Next")
W("    Set blk = doc.Blocks.Add(P3(0#, 0#), pBlkTable)")
W("    If Err.Number <> 0 Then Err.Clear: Exit Sub")
W("    On Error GoTo 0")
W("")
W("    ' --- 외곽 (닫힌 폴리라인) ---")
W("    Dim a0(0 To %d) As Double"%(len(poly)*2-1))
for i,(x,y) in enumerate(poly):
    W("    a0(%d) = %s: a0(%d) = %s"%(i*2,fmt(x),i*2+1,fmt(y)))
W("    Set e = blk.AddLightWeightPolyline(a0)")
W("    e.Closed = True")
W("")
W("    ' --- 내부 구분선 %d개 ---"%len(lines))
for (x1,y1,x2,y2) in lines:
    W("    blk.AddLine P3(%s, %s), P3(%s, %s)"%(fmt(x1),fmt(y1),fmt(x2),fmt(y2)))
W("")
W("    ' --- 고정 문자 %d개 ---"%len(texts))
for (s,x,y,h,ha,va,ap) in texts:
    s2=s.replace('"','""')
    W("    Set e = blk.AddText(\"%s\", P3(%s, %s), %s)"%(s2,fmt(x),fmt(y),fmt(h)))
    av=A(ha,va)
    if av!=0:
        W("    e.Alignment = %d: e.TextAlignmentPoint = P3(%s, %s)"%(av,fmt(ap[0]),fmt(ap[1])))
W("")
W("    ' --- 속성 정의 %d개 ---"%len(atts))
for (tag,x,y,h,ha,va) in atts:
    W("    Set at = blk.AddAttribute(%s, 0, \"%s\", P3(%s, %s), \"%s\", \"\")"%(fmt(h),tag,fmt(x),fmt(y),tag))
    av=A(ha,va)
    if av!=0:
        W("    at.Alignment = %d: at.TextAlignmentPoint = P3(%s, %s)"%(av,fmt(x),fmt(y)))
W("End Sub")
W("")

# ---------------- 도면번호 ----------------
b=blk('도면번호')
atts=[(e.dxf.tag,attr_of(e)[0],attr_of(e)[1],e.dxf.height,e.dxf.halign,e.dxf.valign)
      for e in b if e.dxftype()=='ATTDEF']
W("' 도면번호(표제란) : 속성 %d개 (원본은 도곽 XREF 위에 이 속성만 얹는다)"%len(atts))
W("Public Sub CreateTitleBlock()")
W("    Dim blk As Object, at As Object")
W("    If BlockExists(pBlkTitle) Then Exit Sub")
W("    On Error Resume Next")
W("    Set blk = doc.Blocks.Add(P3(0#, 0#), pBlkTitle)")
W("    If Err.Number <> 0 Then Err.Clear: Exit Sub")
W("    On Error GoTo 0")
for (tag,x,y,h,ha,va) in atts:
    W("    Set at = blk.AddAttribute(%s, 0, \"%s\", P3(%s, %s), \"%s\", \"\")"%(fmt(h),tag,fmt(x),fmt(y),tag))
    av=A(ha,va)
    if av!=0:
        W("    at.Alignment = %d: at.TextAlignmentPoint = P3(%s, %s)"%(av,fmt(x),fmt(y)))
W("End Sub")
W("")

# ---------------- HDot ----------------
b=blk('HDot')
W("' HDot : 치수 보조점 (폭 1.0 의 짧은 선, 삽입축척 0.02)")
W("Public Sub CreateDotBlock()")
W("    Dim blk As Object, e As Object")
W("    Dim a0(0 To 3) As Double")
W("    If BlockExists(pBlkDot) Then Exit Sub")
W("    If Len(pBlkDot) = 0 Then Exit Sub")
W("    On Error Resume Next")
W("    Set blk = doc.Blocks.Add(P3(0#, 0#), pBlkDot)")
W("    If Err.Number <> 0 Then Err.Clear: Exit Sub")
W("    a0(0) = 0.5: a0(1) = 0#: a0(2) = -0.5: a0(3) = 0#")
W("    Set e = blk.AddLightWeightPolyline(a0)")
W("    e.ConstantWidth = 1#")
W("    On Error GoTo 0")
W("End Sub")
W("")
open('/home/user/webapp/_src/_blocks_part.bas','w',encoding='utf-8').write('\n'.join(out)+'\n')
print('\n'.join(out[:12]))
print('TOTAL LINES', len(out))
