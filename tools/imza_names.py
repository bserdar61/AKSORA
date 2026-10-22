# final xlsm'e imza resmi base64 parcalarini gizli tanimli ad olarak ekler (yoksa).
import zipfile,sys,shutil,os,re
src=sys.argv[1]; b64=open(os.path.join(os.path.dirname(__file__),'imza_b64.txt')).read().strip()
z=zipfile.ZipFile(src); wb=z.read('xl/workbook.xml').decode('utf-8')
parca=[b64[i:i+200] for i in range(0,len(b64),200)]
wb=re.sub(r'<definedName name="_AKS_IMZA_\d+"[^>]*>[^<]*</definedName>','',wb)
yeni=['<definedName name="_AKS_IMZA_%03d" hidden="1">"%s"</definedName>'%(i+1,p) for i,p in enumerate(parca)]
m=re.search(r'<definedNames>(.*?)</definedNames>',wb,re.S)
var=re.findall(r'<definedName\b.*?</definedName>',m.group(1),re.S) if m else []
hepsi=var+yeni
blok='<definedNames>'+''.join(hepsi)+'</definedNames>'
wb=wb[:m.start()]+blok+wb[m.end():] if m else wb.replace('</sheets>','</sheets>'+blok,1)
tmp=src+'.tmp'
with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as o:
    for it in z.infolist():
        data=z.read(it.filename)
        if it.filename=='xl/workbook.xml': data=wb.encode('utf-8')
        o.writestr(it,data)
z.close(); shutil.move(tmp,src); print('imza adlari',len(parca))
