import re,sys
def clean(s):
    out,inq=[],False
    for c in s:
        if c=='"': inq=not inq; out.append(' ')
        elif c=="'" and not inq: break
        else: out.append(' ' if inq else c)
    return ''.join(out)
def merge(raw):
    logical,i=[],0
    while i<len(raw):
        st,buf,cl=i+1,raw[i],clean(raw[i])
        while re.search(r'\s_\s*$',cl) and i+1<len(raw):
            buf=re.sub(r'\s_\s*$',' ',buf); i+=1; buf+=raw[i].strip(); cl=clean(buf)
        logical.append((st,i+1,cl)); i+=1
    return logical
KW=set("""and as byref byval call case const dim do each else elseif empty end enum eqv
erase event exit false for friend function get global gosub goto if imp implements in is
let like lock loop lset me mod new next not nothing null on option optional or paramarray
preserve private property public put raiseevent redim rem resume return rset select set
static step stop sub then to true type until wend while with withevents xor
boolean byte currency date decimal double integer long object single string variant any
ptr err
""".split())
PREFIX=re.compile(r'^\s*(?:WithEvents|Static|Optional|ByVal|ByRef|ParamArray|Preserve|Public|Private|Global|Dim|Friend)\s+',re.I)
ID=re.compile(r'([A-Za-z_\uac00-\ud7a3][\w]*)')
def split_top(s):
    parts,d,cur=[],0,''
    for ch in s:
        if ch in '([': d+=1
        elif ch in ')]': d-=1
        if ch==',' and d==0: parts.append(cur); cur=''
        else: cur+=ch
    parts.append(cur); return parts
def names_of(t):
    out=[]
    for p in split_top(t):
        for _ in range(8):
            n=PREFIX.sub('',p,count=1)
            if n==p: break
            p=n
        m=ID.match(p.strip())
        if m and m.group(1).lower() not in KW: out.append(m.group(1))
    return out
SIG=re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_\uac00-\ud7a3][\w]*)\s*(\()?',re.I)
ENDS=re.compile(r'^\s*End\s+(Sub|Function|Property)\b',re.I)
BUILTIN=set("""err msgbox format left right mid trim lcase ucase len cstr clng cint cdbl csng cbool
cdate isnumeric isempty isnull isarray isobject iserror instr instrrev replace split join array
ubound lbound abs int fix sgn sqr sin cos tan atn exp log round rnd val str chr asc hex oct
now date time timer doevents appactivate sendkeys createobject getobject callbyname evaluate
typename vartype iif switch choose rgb space string strcomp strconv thisworkbook application
worksheets sheets workbooks error debug print""".split())
HOST=set('acad doc ms cadappname'.split())
def run(path,label):
    raw=open(path,encoding='utf-8',errors='replace').read().split('\n')
    L=merge(raw)
    procs,i=[],0
    while i<len(L):
        t=L[i][2]; m=SIG.match(t)
        if m and not ENDS.match(t) and not re.match(r'^\s*(End|Exit)\b',t,re.I):
            end=len(L)-1
            for j in range(i+1,len(L)):
                if ENDS.match(L[j][2]): end=j; break
            procs.append(dict(name=m.group(2),sig=i,end=end,line=L[i][0],sigtext=t)); i=end+1
        else: i+=1
    inb=[False]*len(L)
    for p in procs:
        for j in range(p['sig'],p['end']+1): inb[j]=True
    module=set(); intype=False
    for j,(s,e,t) in enumerate(L):
        if inb[j]: continue
        st=t.strip()
        if not st: continue
        if re.match(r'^(?:Public\s+|Private\s+)?Type\s+',st,re.I): intype=True; continue
        if intype:
            if re.match(r'^End\s+Type',st,re.I): intype=False
            continue
        if re.match(r'^(?:Public|Private|Global|Friend)\s+Declare\b',st,re.I): continue
        m=re.match(r'^(?:(?:Public|Private|Global|Friend|Dim)\s+)?Const\s+(.+)$',st,re.I)
        if m:
            for n in names_of(m.group(1)): module.add(n.lower())
            continue
        m=re.match(r'^(?:Public|Private|Global|Dim|Friend)\s+(.+)$',st,re.I)
        if m:
            for n in names_of(m.group(1)): module.add(n.lower())
    types=set()
    for (s,e,t) in L:
        for m in re.finditer(r'\bAs\s+(?:New\s+)?([A-Za-z_\uac00-\ud7a3][\w]*)',t,re.I):
            types.add(m.group(1).lower())
    procnames=set(p['name'].lower() for p in procs)
    report=[]
    for p in procs:
        local=set()
        for j in range(p['sig']+1,p['end']):
            st=L[j][2].strip()
            m=re.match(r'^(?:Dim|Static|Const)\s+(.+)$',st,re.I)
            if m:
                for n in names_of(m.group(1)): local.add(n.lower())
                continue
            m=re.match(r'^ReDim\s+(?:Preserve\s+)?(.+)$',st,re.I)
            if m:
                for n in names_of(m.group(1)): local.add(n.lower())
        params=set()
        sig=p['sigtext']
        if '(' in sig:
            d,body=0,''
            for k in range(sig.index('('),len(sig)):
                ch=sig[k]
                if ch=='(':
                    d+=1
                    if d==1: continue
                elif ch==')':
                    d-=1
                    if d==0: break
                body+=ch
            for n in names_of(body): params.add(n.lower())
        known=local|params|module|procnames|BUILTIN|KW|HOST|types
        bad=[]
        for j in range(p['sig']+1,p['end']):
            s0,e0,t=L[j]
            st=t.strip()
            if not st: continue
            w=t
            w=re.sub(r'\bOn\s+Error\b.*$','',w,flags=re.I)
            w=re.sub(r'\bAs\s+(?:New\s+)?[A-Za-z_\uac00-\ud7a3][\w]*','',w,flags=re.I)
            w=re.sub(r'[.!][A-Za-z_\uac00-\ud7a3][\w]*','',w)
            w=re.sub(r'([A-Za-z_\uac00-\ud7a3][\w]*)\s*:=','',w)
            w=re.sub(r'\b(GoTo|GoSub)\s+[A-Za-z_\uac00-\ud7a3][\w]*','',w,flags=re.I)
            w=re.sub(r'^\s*[A-Za-z_\uac00-\ud7a3][\w]*:\s*$','',w)
            w=clean(w)
            for m in re.finditer(r'[A-Za-z_\uac00-\ud7a3][\w]*',w):
                nm=m.group(0); lo=nm.lower()
                if lo in known or lo in KW: continue
                if re.match(r'^(vb|xl|ac|mso|wiz|ofc)[A-Z]',nm): continue
                if re.match(r'^[Rr]\d',nm): continue
                if re.match(r'^[0-9]',nm): continue
                bad.append((nm,s0))
        if bad:
            u={}
            for nm,ln in bad: u.setdefault(nm,ln)
            report.append((p['name'],p['line'],sorted(u.items(),key=lambda x:x[1])))
    print('=== %s ===  프로시저 %d / 모듈선언 %d / 타입명 %d'%(label,len(procs),len(module),len(types)))
    if not report: print('  OK 미선언 없음')
    for nm,ln,items in report:
        print('  [%s] L%d'%(nm,ln))
        for a,b in items: print('      %-22s L%d'%(a,b))
    print()
    return report
run('_src/횡단면도그리기_utf8.bas','UTF-8')
run('_src/횡단면도그리기.bas','CP949')
