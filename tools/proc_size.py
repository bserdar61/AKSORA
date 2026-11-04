# Risk siniri: tek prosedur kaynagi 48 KB'i (VBA "Procedure too large" 64 KB derlenmis siniri), modul 600 KB'i asmasin.
import re,glob
P=re.compile(r'^(?:Private |Public |Friend )?(?:Static )?(Sub|Function|Property \w+) +(\w+)',re.I)
bad=0
for f in glob.glob('raw/*.txt'):
    t=open(f,encoding='cp1254',newline='').read()
    if len(t)>600000: print('MODUL BUYUK',f,len(t)); bad+=1
    cur=None
    for l in t.split('\r\n'):
        m=P.match(l)
        if m: cur=m.group(2); sz=0
        if cur: sz+=len(l)+2
        if cur and re.match(r'^End (Sub|Function|Property)',l,re.I):
            if sz>48000: print('PROSEDUR BUYUK',f,cur,sz); bad+=1
            cur=None
print('boyut sorun',bad)
