# Moduller arasi Private erisim denetimi: bir modulde Private tanimli ad, baska modulde (kendi tanimi ve Public karsiligi yokken) kullaniliyor mu.
import re,glob
P=r'^(?:(Private|Public|Friend) )?(?:Static )?(?:Sub|Function|Property (?:Get|Let|Set)) +(\w+)'
mods={}
for f in glob.glob('raw/*.txt'):
    t=open(f,encoding='cp1254',newline='').read().split('\r\n'); mods[f.split('/')[-1][:-4]]=t
priv={};pub=set();own={}
for m,t in mods.items():
    own[m]=set(); cls = m.startswith(('Sayfa','BuÇal','frm')) or 'BuÇal' in m
    for l in t:
        x=re.match(P,l)
        if x:
            n=x.group(2); own[m].add(n.lower())
            if x.group(1)=='Private': priv.setdefault(n.lower(),set()).add(m)
            elif not cls: pub.add(n.lower())
        x=re.match(r'^(Private|Public|Dim|Global)\s+(?:Const\s+|Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+|Type\s+|Enum\s+)?(\w+)',l)
        if x and not re.match(P,l):
            n=x.group(2).lower(); own[m].add(n)
            if x.group(1) in('Public','Global') and not cls: pub.add(n)
            elif x.group(1) in ('Private','Dim'): priv.setdefault(n,set()).add(m)
bad=0
for m,t in mods.items():
    toks=set()
    for l in t:
        s=re.sub(r'"[^"]*"','""',l)
        if "'" in s: s=s[:s.index("'")]
        toks|={w.lower() for w in re.findall(r'(?<![.\w])[A-Za-z_]\w*',s)}
    for n in toks:
        if n in priv and n not in own[m] and n not in pub and not (priv[n]-{m})=={*()} :
            # yerel degisken olabilir: modulde Dim/parametre olarak geciyor mu
            body='\n'.join(t)
            if re.search(r'(?i)\b(Dim|ByVal|ByRef|Optional|Static|ReDim)\b[^\n]*\b'+re.escape(n)+r'\b',body) or re.search(r'(?i)\(\s*'+re.escape(n)+r'\s+As\b',body) or re.search(r'(?i),\s*'+re.escape(n)+r'\s+As\b',body): continue
            print('PRIVATE ERISIM?',m,'->',n,sorted(priv[n])); bad+=1
print('xmod sorun',bad)
