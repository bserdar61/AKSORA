import re,glob,sys
bad=0
for f in sorted(glob.glob('raw/*.txt')):
    t=open(f,encoding='cp1254').read().replace('\r\n','\n').split('\n')
    ilk=None; derin=0
    for i,l in enumerate(t):
        s=l.strip()
        if re.match(r"^(Public |Private |Friend )?(Static )?(Sub|Function|Property (Get|Let|Set)) ",s):
            ilk=ilk if ilk is not None else i
            if not re.search(r'End (Sub|Function|Property)\s*$',s): derin+=1
        elif re.match(r'^End (Sub|Function|Property)\b',s): derin-=1
        elif ilk is not None and derin==0 and re.match(r'^(Public|Private|Dim|Const|Global)\s',s) and not re.match(r'^(Public |Private )?(Declare|Type|Enum)\b',s):
            print('MODUL BILDIRIMI PROSEDURDEN SONRA:',f,i+1,s[:90]); bad+=1
print('decl errors',bad)
