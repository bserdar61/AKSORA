# Ayni prosedurde ayni adla (buyuk/kucuk harf farkiyla) iki kez tanimlanan degisken / parametre: VBA derleme hatasi.
import re,glob
P=re.compile(r'^(?:Private |Public |Friend )?(?:Static )?(?:Sub|Function|Property \w+) +(\w+)\((.*)\)',re.I)
bad=0
for f in glob.glob('raw/*.txt'):
    L=open(f,encoding='cp1254',newline='').read().split('\r\n')
    cur=None; buf=''; els=False
    for i,l in enumerate(L):
        if cur is None:
            if buf or P.match(l.split("'")[0]) or re.match(r'^(?:Private |Public |Friend )?(?:Sub|Function|Property \w+) ',l):
                buf+=l.rstrip(' _')+' '
                if l.rstrip().endswith(' _'): continue
                m=P.match(buf); buf=''
                if not m: continue
                cur=m.group(1); names={}
                for v in re.findall(r'(?:^|,)\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+|ParamArray\s+)?(\w+)\s*(?:\(\))?\s+As',m.group(2)):
                    names[v.lower()]=i
            continue
        s=re.sub(r'"[^"]*"','""',l).split("'")[0]
        if re.match(r'\s*#Else',s): els=True
        if re.match(r'\s*#End If',s): els=False
        if re.match(r'\s*(Dim|Static) ',s) and not els:
            for v in re.findall(r'(?:(?:Dim|Static)\s+|,\s*)(\w+)\s*(?:\([^)]*\))?\s+As',s):
                if v.lower() in names: print('CIFT TANIM',f,cur,v,i+1); bad+=1
                names[v.lower()]=i
        if re.match(r'^End (Sub|Function|Property)',l): cur=None
print('cift tanim',bad)
