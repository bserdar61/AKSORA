# Var olan vbaProject.bin'e yeni standart modul kaydi ekler (dir + PROJECT + PROJECTwm). Kaynak build_bin ile raw'dan gelir.
import sys,os,struct; sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import olefile, ovba, vbadir, cfb
src, out, ad, sablon = sys.argv[1:5]
o=olefile.OleFileIO(src)
recs=vbadir.parse(ovba.decompress(o.openstream('VBA/dir').read()))
names=[dt.decode('cp1254') for r,dt in recs if r==0x19]
assert ad not in names, 'modul zaten var'
i=next(k for k,(r,dt) in enumerate(recs) if r==0x19 and dt.decode('cp1254')==sablon)
j=i
while recs[j][0]!=0x2B: j+=1
blok=[[r,dt] for r,dt in recs[i:j+1]]
ab=ad.encode('cp1254'); au=ad.encode('utf-16le')
for b in blok:
    if b[0] in (0x19,0x1A): b[1]=ab
    elif b[0] in (0x47,0x32): b[1]=au
    elif b[0]==0x31: b[1]=struct.pack('<I',0)
    elif b[0]==0x2C: b[1]=struct.pack('<H',0xFFFF)
assert any(b[0]==0x21 for b in blok), 'sablon standart modul degil'
son=max(k for k,(r,dt) in enumerate(recs) if r==0x2B)
recs=recs[:son+1]+blok+recs[son+1:]
k=next(k for k,(r,dt) in enumerate(recs) if r==0x0F)
recs[k][1]=struct.pack('<H',struct.unpack('<H',recs[k][1])[0]+1)
newdir=vbadir.build(recs)
p=o.openstream('PROJECT').read()
lines=p.split(b'\r\n')
last=max(n for n,l in enumerate(lines) if l.startswith(b'Module='))
lines.insert(last+1,b'Module='+ab)
newp=b'\r\n'.join(lines)
w=o.openstream('PROJECTwm').read()
assert w.endswith(b'\x00\x00')
neww=w[:-2]+ab+b'\x00'+au+b'\x00\x00'+b'\x00\x00'
def clsid_of(e):
    if not e.clsid: return b'\0'*16
    import uuid; return uuid.UUID(e.clsid).bytes_le
def ent(path): return o.direntries[o._find(path)]
root=cfb.Node('Root Entry',5); root.clsid=clsid_of(o.direntries[0])
root.children.append(cfb.Node('PROJECT',2,newp)); root.children.append(cfb.Node('PROJECTwm',2,neww))
vba=cfb.Node('VBA',1,clsid=clsid_of(ent('VBA'))); root.children.append(vba)
vba.children.append(cfb.Node('dir',2,ovba.compress(newdir)))
# Modul akislari yer tutucu (build_bin hepsini raw kaynagindan yeniden yazar); p-code / SRP alinmaz.
vba.children.append(cfb.Node('_VBA_PROJECT',2,bytes.fromhex('cc61ffff000300')))
for m in vbadir.modules(recs):
    vba.children.append(cfb.Node(m['streamu'],2,ovba.compress(('Attribute VB_Name = "'+m['name']+'"\r\n').encode('cp1254'))))
for e in o.listdir(streams=False,storages=True):
    if len(e)==1 and e[0]!='VBA':
        st=cfb.Node(e[0],1,clsid=clsid_of(ent(e[0]))); root.children.append(st)
        for s in o.listdir():
            if len(s)==2 and s[0]==e[0]: st.children.append(cfb.Node(s[1],2,o.openstream(s).read()))
cfb.write(root,out); print('eklendi',ad,'modul sayisi',len(names)+1)
