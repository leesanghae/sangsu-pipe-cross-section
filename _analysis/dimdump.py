#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""k=0 단면의 LEAD/LEAT/DIML 엔티티를 셀 원점 기준 상대좌표로 덤프"""
import ezdxf
d = ezdxf.readfile('/home/user/uploaded_files/종평면도(상수관로) (횡단포함).txt')
msp = d.modelspace()

# 셀 원점(k=0) : 스펙에서 확인된 값
import sys
sys.path.insert(0, '/home/user/webapp/_analysis')
from calc import load
secs, _wb = load()
s0 = secs[0]
print("k=0 : sta=%s GH=%s INV=%s D=%s code=%s" % (
    s0['sta'], s0['gh'], s0['inv'], s0['dia'], s0['code']))

# 원점 찾기 : CC-TABL 블록 삽입점 - pTblDX(5.0)
tabl = [e for e in msp if e.dxftype() == 'INSERT' and e.dxf.name == '관로06']
print("관로06 삽입 %d개" % len(tabl))
ins = sorted(tabl, key=lambda e: (-e.dxf.insert[1], e.dxf.insert[0]))
e0 = ins[0]
cx = e0.dxf.insert[0] - 5.0
cy = e0.dxf.insert[1]
print("k=0 셀 원점 = (%.4f, %.4f)" % (cx, cy))

GH = s0['gh']
H = s0['gh'] - s0['inv']

def rel(x, y):
    return (x - cx, y - cy)

print("\n--- LEAD (보조선) ---")
for e in msp:
    if e.dxftype() != 'LINE' or e.dxf.layer != 'CC-LEAD':
        continue
    if not (abs(e.dxf.start[0]-cx) < 9 and abs(e.dxf.start[1]-cy) < 9):
        continue
    a = rel(e.dxf.start[0], e.dxf.start[1])
    b = rel(e.dxf.end[0], e.dxf.end[1])
    if abs(a[0]-b[0]) > 25 or abs(a[1]-b[1]) > 25:
        continue
    print("  (%.3f,%.3f) -> (%.3f,%.3f)   | dy=%+.2f-H=%+.3f"
          % (a[0], a[1], b[0], b[1], a[1], a[1]+H))

print("\n--- LEAT (문자) ---")
for e in msp:
    if e.dxf.layer != 'CC-LEAT':
        continue
    p = None
    try:
        p = e.dxf.align_point
    except Exception:
        p = None
    if p is None:
        p = e.dxf.insert
    a = rel(p[0], p[1])
    if abs(a[0]) > 12 or abs(a[1]) > 12:
        continue
    print("  '%-14s' (%.3f,%.3f) h=%.2f | y-H=%+.3f" % (e.dxf.text, a[0], a[1], e.dxf.height, a[1]+H))
