import sys
def rd(p): return open(p,encoding='cp1254').read().replace('\r\n','\n')
def wr(p,s):
    b=s.replace('\n','\r\n').encode('cp1254')   # once kodla: hata olursa dosya bozulmaz
    open(p,'wb').write(b)
class F:
    def __init__(s,p): s.p='raw/'+p+'.txt'; s.s=rd(s.p)
    def R(s,o,n,c=1):
        k=s.s.count(o); assert k==c,(s.p,k,o[:100]); s.s=s.s.replace(o,n)
    def save(s): wr(s.p,s.s)
