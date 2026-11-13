import re
src = open('/home/user/webapp/_src/횡단면도그리기_utf8.bas', encoding='utf-8').read().splitlines()

# 첫 프로시저(Sub/Function/Property) 시작 줄
first_proc = None
for i, ln in enumerate(src, 1):
    if re.match(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property (Get|Let|Set))\b', ln):
        first_proc = i
        break
print("첫 프로시저 시작 줄:", first_proc)

# 모듈 레벨 선언 후보 (들여쓰기 없음)
decl_re = re.compile(r'^(Public |Private |Friend )?(Const|Dim|Type|Enum|Declare)\b|^(Type|Enum)\s+\w+\s*$')
bad = []
for i, ln in enumerate(src, 1):
    if decl_re.match(ln):
        if first_proc and i > first_proc:
            bad.append((i, ln.strip()))
print()
if bad:
    print("*** 프로시저 뒤에 있는 모듈레벨 선언 (VBA 위반) ***  %d 건" % len(bad))
    for i, ln in bad:
        print(f"   {i:5d}: {ln}")
else:
    print("프로시저 뒤 모듈레벨 선언: 없음")
