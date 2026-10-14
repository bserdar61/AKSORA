import re,sys,glob,os
files=sys.argv[1:]
KW=set('''if then else elseif end sub function for next to step each in do loop while until wend with select case is not and or xor mod dim as new set let call exit goto on error resume private public static const true false nothing empty null me byval byref optional paramarray property get option explicit integer long double single string boolean variant object date currency byte decimal redim preserve erase debug print and_also typeof like eqv imp lbound ubound len''' .split())
allpub=set()
alltext={}
for f in glob.glob(os.path.dirname(files[0])+'/*.txt'):
    t=open(f,'rb').read().decode('cp1254'); alltext[f]=t
    for m in re.finditer(r'(?im)^\s*(?:public\s+|global\s+)?(?:sub|function|property\s+get|const|declare\s+ptrsafe\s+sub|declare\s+sub)?\s*(\w+)',t): pass
    for m in re.finditer(r'(?im)^(?:Public\s+|Friend\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)',t): allpub.add(m.group(1).lower())
    for m in re.finditer(r'(?im)^\s*(?:Public|Global|Private|Dim)\s+(?!Sub|Function|Property|Declare|Const)(.+)$',t):
        if not re.match(r'(?i)\s*(Sub|Function)',m.group(1)):
            for v in m.group(1).split(','):
                n=re.match(r'\s*(?:WithEvents\s+)?(\w+)',v)
                if n: allpub.add(n.group(1).lower())
    for m in re.finditer(r'(?im)^\s*(?:Public\s+|Private\s+)?Const\s+(\w+)',t): allpub.add(m.group(1).lower())
for f in files:
    t=open(f,'rb').read().decode('cp1254')
    # join continuations
    t=re.sub(r' _\r?\n',' ',t)
    lines=t.split('\r\n') if '\r\n' in t else t.split('\n')
    stack=[]
    proc=None; decl=set()
    for no,l in enumerate(lines,1):
        s=re.sub(r'"[^"]*"','""',l)
        s=re.sub(r"'.*$",'',s).strip()
        if not s or s.lower().startswith('attribute'): continue
        for part in re.split(r':(?!=)',s):
            p=part.strip(); pl=p.lower()
            if not p: continue
            if re.match(r'^\w+$',p) and not pl in KW and part is not None and p==s.strip(): 
                pass
            m=re.match(r'(?i)^(?:public\s+|private\s+|friend\s+)?(?:static\s+)?(sub|function|property\s+(?:get|let|set))\s+(\w+)\s*\((.*)\)',p)
            if m:
                stack.append(('proc',no)); proc=m.group(2); decl=set()
                for prm in m.group(3).split(','):
                    n=re.match(r'\s*(?:optional\s+)?(?:byval\s+|byref\s+)?(?:paramarray\s+)?(\w+)',prm,re.I)
                    if n: decl.add(n.group(1).lower())
                decl.add(proc.lower()); continue
            if re.match(r'(?i)^end\s+(sub|function|property)$',p):
                if not stack or stack[-1][0]!='proc': print(f,no,'unbalanced end proc',stack[-3:])
                else: stack.pop()
                proc=None; continue
            if re.match(r'(?i)^(dim|static|redim(\s+preserve)?)\s',p):
                body=re.sub(r'(?i)^(dim|static|redim(\s+preserve)?)\s+','',p)
                body=re.sub(r'\([^()]*\)','',body)
                for v in body.split(','):
                    n=re.match(r'\s*(\w+)',v)
                    if n: decl.add(n.group(1).lower())
                continue
            if re.match(r'(?i)^if\b.*\bthen$',p): stack.append(('if',no)); continue
            if re.match(r'(?i)^if\b.*\bthen\s+\S',p): continue  # single-line if
            if re.match(r'(?i)^end\s+if$',p):
                if not stack or stack[-1][0]!='if': print(f,no,'unbalanced end if',stack[-3:])
                else: stack.pop()
                continue
            if re.match(r'(?i)^(for)\b',p): stack.append(('for',no)); continue
            if re.match(r'(?i)^next\b',p):
                if not stack or stack[-1][0]!='for': print(f,no,'unbalanced next',stack[-3:])
                else: stack.pop()
                continue
            if re.match(r'(?i)^with\b',p): stack.append(('with',no)); continue
            if re.match(r'(?i)^end\s+with$',p):
                if not stack or stack[-1][0]!='with': print(f,no,'unbalanced end with',stack[-3:])
                else: stack.pop()
                continue
            if re.match(r'(?i)^do\b',p): stack.append(('do',no)); continue
            if re.match(r'(?i)^loop\b',p):
                if not stack or stack[-1][0]!='do': print(f,no,'unbalanced loop',stack[-3:])
                else: stack.pop()
                continue
            if re.match(r'(?i)^select\s+case\b',p): stack.append(('sel',no)); continue
            if re.match(r'(?i)^end\s+select$',p):
                if not stack or stack[-1][0]!='sel': print(f,no,'unbalanced end select',stack[-3:])
                else: stack.pop()
                continue
            # assignment target check
            if proc:
                m=re.match(r'(?i)^(?:set\s+)?(\w+)\s*(?:\([^=]*\))?\s*=',p)
                if m:
                    n=m.group(1).lower()
                    if n not in decl and n not in allpub and n not in KW: print(f,no,'UNDECLARED?',m.group(1),'in',proc)
                m=re.match(r'(?i)^for\s+(?:each\s+)?(\w+)',p)
    if stack: print(f,'unclosed',stack)

# module-level declarations after the first procedure are a compile error in VBA
import glob as _g
for _f in sys.argv[1:]:
    _t=open(_f,'rb').read().decode('cp1254').split('\r\n')
    _first=None
    for _i,_l in enumerate(_t):
        if re.match(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property) ',_l): _first=_i; break
    if _first is None: continue
    for _i,_l in enumerate(_t[_first:],_first):
        if re.match(r'^(Private|Public|Global)\s+(Const\s+|Declare\s+|Type\s+|Enum\s+|WithEvents\s+)?\w+',_l) and not re.match(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property) ',_l):
            print(_f,_i+1,'MODULE-LEVEL AFTER PROCEDURE:',_l)
print('lint done')
