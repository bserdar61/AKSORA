# AKSORA VBA derleme araclari.
#   python3 tools/vbakod.py cikar <xlsm> <klasor>          : modulleri <klasor>/<ad>.txt olarak cikarir (cp1254, CRLF)
#   python3 tools/vbakod.py derle <temel xlsm> <klasor> <cikti xlsm> : klasordeki .txt kaynaklarla vbaProject.bin'i yeniden yazar
# Derlemede: modul offset 0, _VBA_PROJECT surum 0xFFFF (Excel acilista kaynaktan derler), __SRP_ onbellekleri silinir.
import sys, os, zipfile, io, struct, tempfile, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import olefile, ovba, vbadir, cfb


def _bin(xlsm):
    with zipfile.ZipFile(xlsm) as z:
        return z.read('xl/vbaProject.bin')


def _dosya_adi(ad):
    return ad


def cikar(xlsm, klasor):
    os.makedirs(klasor, exist_ok=True)
    b = _bin(xlsm)
    o = olefile.OleFileIO(io.BytesIO(b))
    recs = vbadir.parse(ovba.decompress(o.openstream('VBA/dir').read()))
    mods = vbadir.modules(recs)
    for m in mods:
        ofs = struct.unpack('<I', recs[m['ofs']][1])[0]
        raw = o.openstream('VBA/' + m['akis']).read()
        src = ovba.decompress(raw[ofs:])
        with open(os.path.join(klasor, _dosya_adi(m['ad']) + '.txt'), 'wb') as f:
            f.write(src)
    o.close()
    return [m['ad'] for m in mods]


def derle(temel, klasor, cikti):
    b = _bin(temel)
    tmp = tempfile.mkdtemp()
    p = os.path.join(tmp, 'in.bin')
    open(p, 'wb').write(b)
    kok = cfb.oku(p)
    vba = next(c for c in kok['cocuk'] if c['ad'] == 'VBA')
    akis = {c['ad']: c for c in vba['cocuk']}
    recs = vbadir.parse(ovba.decompress(akis['dir']['veri']))
    mods = vbadir.modules(recs)
    eksik = []
    for m in mods:
        yol = os.path.join(klasor, m['ad'] + '.txt')
        if not os.path.exists(yol):
            eksik.append(m['ad']); continue
        src = open(yol, 'rb').read()
        akis[m['akis']]['veri'] = ovba.compress(src)
        recs[m['ofs']][1] = struct.pack('<I', 0)
    if eksik:
        raise SystemExit('kaynak yok: ' + ', '.join(eksik))
    akis['dir']['veri'] = ovba.compress(vbadir.build(recs))
    akis['_VBA_PROJECT']['veri'] = struct.pack('<HHBH', 0x61CC, 0xFFFF, 0, 3)
    vba['cocuk'] = [c for c in vba['cocuk'] if not c['ad'].startswith('__SRP_')]
    yeni = os.path.join(tmp, 'vbaProject.bin')
    cfb.write(kok, yeni)
    yb = open(yeni, 'rb').read()
    with zipfile.ZipFile(temel) as zi, zipfile.ZipFile(cikti + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            d = yb if it.filename == 'xl/vbaProject.bin' else zi.read(it.filename)
            zo.writestr(it, d, compress_type=zipfile.ZIP_DEFLATED)
    os.replace(cikti + '.tmp', cikti)
    shutil.rmtree(tmp)
    # dogrulama: geri cikar ve karsilastir
    kt = tempfile.mkdtemp()
    cikar(cikti, kt)
    for m in mods:
        a = open(os.path.join(klasor, m['ad'] + '.txt'), 'rb').read()
        c = open(os.path.join(kt, m['ad'] + '.txt'), 'rb').read()
        assert a == c, 'geri okuma farki: ' + m['ad']
    shutil.rmtree(kt)
    print('derlendi:', cikti, len(mods), 'modul, geri okuma ayni')


if __name__ == '__main__':
    if sys.argv[1] == 'cikar':
        print(len(cikar(sys.argv[2], sys.argv[3])), 'modul')
    elif sys.argv[1] == 'derle':
        derle(sys.argv[2], sys.argv[3], sys.argv[4])
