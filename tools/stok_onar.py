import zipfile, sys, re
from lxml import etree
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
N='{%s}'%NS
iyi, bozuk, cikti = sys.argv[1:4]
zi=zipfile.ZipFile(iyi); zb=zipfile.ZipFile(bozuk)
sst=etree.fromstring(zi.read('xl/sharedStrings.xml'))
ss=[''.join(si.itertext()) for si in sst.findall(N+'si')]
def col(ref): return re.match(r'[A-Z]+',ref).group(0)
def rown(ref): return int(re.search(r'\d+',ref).group(0))
yeni={}
rapor=[]
for sh in ['xl/worksheets/sheet14.xml','xl/worksheets/sheet16.xml']:
    a=etree.fromstring(zi.read(sh)); b=etree.fromstring(zb.read(sh))
    ra={int(r.get('r')):r for r in a.find(N+'sheetData').findall(N+'row')}
    rb={int(r.get('r')):r for r in b.find(N+'sheetData').findall(N+'row')}
    son=max(r for r,row in ra.items() if r>=7 and any(c.find(N+'v') is not None for c in row if col(c.get('r'))=='G'))
    deg=0; fsay=0
    for r in range(7,son+1):
        if r not in ra or r not in rb: continue
        ca={col(c.get('r')):c for c in ra[r]}; cb={col(c.get('r')):c for c in rb[r]}
        for L in 'ABCDEFGHIJKL':
            x=ca.get(L); y=cb.get(L)
            if y is None: continue
            if x is not None and x.find(N+'f') is not None: fsay+=1
            # deger
            val=None; tip=None
            if x is not None and x.find(N+'v') is not None:
                v=x.find(N+'v').text
                if x.get('t')=='s': val=ss[int(v)]; tip='str'
                elif x.get('t')=='str': val=v; tip='str'
                elif x.get('t')=='b': val=v; tip='b'
                else: val=v; tip='n'
            elif x is not None and x.get('t')=='inlineStr':
                val=''.join(x.itertext()); tip='str'
            s=y.get('s')
            for ch in list(y): y.remove(ch)
            for k in list(y.attrib):
                if k not in ('r','s'): del y.attrib[k]
            if val is None: continue
            if tip=='str':
                y.set('t','inlineStr'); is_=etree.SubElement(y,N+'is'); t=etree.SubElement(is_,N+'t'); t.text=val
                if val!=val.strip(): t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            else:
                if tip=='b': y.set('t','b')
                etree.SubElement(y,N+'v').text=val
            deg+=1
    rapor.append((sh,son,deg,fsay))
    yeni[sh]=etree.tostring(b,xml_declaration=True,encoding='UTF-8',standalone=True)
print(rapor)
with zipfile.ZipFile(cikti,'w',zipfile.ZIP_DEFLATED) as zo:
    for it in zb.infolist():
        d=yeni.get(it.filename) or zb.read(it.filename)
        if it.filename=='xl/calcChain.xml': continue
        zo.writestr(it,d,compress_type=zipfile.ZIP_DEFLATED)
