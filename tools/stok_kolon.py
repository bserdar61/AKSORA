import zipfile,re,sys
src,dst=sys.argv[1:3]
z=zipfile.ZipFile(src)
wbx=z.read('xl/workbook.xml').decode(); rels=z.read('xl/_rels/workbook.xml.rels').decode()
def sf(n):
    rid=re.search(r'<sheet [^>]*name="%s"[^>]*r:id="(rId\d+)"'%re.escape(n),wbx).group(1)
    return 'xl/'+re.search(r'Id="%s"[^>]*Target="([^"]+)"'%rid,rels).group(1)
hedef={1:15,6:19,8:44}
yeni={}
for n in ['CDP_STOK','EPP_STOK','STOKLAR','İLÇE_STOK','YÜKLENİCİ_STOK']:
    f=sf(n); x=z.read(f).decode()
    def rep(m):
        a,b=int(m.group(1)),int(m.group(2))
        if a==b and a in hedef:
            return re.sub(r'width="[\d.]+"','width="%s"'%hedef[a],m.group(0))
        return m.group(0)
    x2=re.sub(r'<col min="(\d+)" max="(\d+)"[^>]*/>',rep,x)
    assert x2!=x, n
    yeni[f]=x2.encode()
with zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as zo:
    for it in z.infolist():
        zo.writestr(it, yeni.get(it.filename) or z.read(it.filename), compress_type=zipfile.ZIP_DEFLATED)
print('ok',len(yeni))
