# Basit CFB (v3, 512 bayt sektor) yazici. Girdi: agac = {'ad': ..., 'clsid': b'..16', 'cocuk': [...], 'veri': bytes|None}
import struct

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
SEC = 512
MINI = 64
CUTOFF = 4096


def _key(ad):
    return (len(ad), ad.upper())


def write(kok, yol):
    # Dizin girisleri duz listesi
    ents = []

    def ekle(node, tip):
        e = {'ad': node['ad'], 'tip': tip, 'clsid': node.get('clsid', b'\x00' * 16), 'veri': node.get('veri'),
             'sol': NOSTREAM, 'sag': NOSTREAM, 'cocuk': NOSTREAM, 'start': ENDOFCHAIN, 'size': 0}
        ents.append(e)
        idx = len(ents) - 1
        if tip in (1, 5):
            kids = []
            for c in node.get('cocuk', []):
                kids.append(ekle(c, 1 if c.get('veri') is None else 2))
            kids.sort(key=lambda k: _key(ents[k]['ad']))

            def agac(lst):
                if not lst:
                    return NOSTREAM
                m = len(lst) // 2
                ents[lst[m]]['sol'] = agac(lst[:m])
                ents[lst[m]]['sag'] = agac(lst[m + 1:])
                return lst[m]
            e['cocuk'] = agac(kids)
        return idx

    ekle(kok, 5)
    # Mini akis
    mini = bytearray()
    minifat = []
    for e in ents:
        if e['tip'] == 2 and e['veri'] is not None and len(e['veri']) < CUTOFF:
            d = e['veri']
            e['size'] = len(d)
            if len(d) == 0:
                e['start'] = ENDOFCHAIN
                continue
            n = (len(d) + MINI - 1) // MINI
            st = len(minifat)
            e['start'] = st
            for k in range(n):
                minifat.append(st + k + 1 if k < n - 1 else ENDOFCHAIN)
            mini += d + b'\x00' * (n * MINI - len(d))
    # Normal sektorler: buyuk akislar, mini akis, minifat, dizin, sonra FAT
    sectors = []  # bayt bloklari (her biri SEC)
    fat = []

    def zincir(data):
        n = (len(data) + SEC - 1) // SEC
        if n == 0:
            return ENDOFCHAIN
        st = len(sectors)
        for k in range(n):
            sectors.append(data[k * SEC:(k + 1) * SEC].ljust(SEC, b'\x00'))
            fat.append(st + k + 1 if k < n - 1 else ENDOFCHAIN)
        return st

    for e in ents:
        if e['tip'] == 2 and e['veri'] is not None and len(e['veri']) >= CUTOFF:
            e['size'] = len(e['veri'])
            e['start'] = zincir(e['veri'])
    ents[0]['start'] = zincir(bytes(mini)) if mini else ENDOFCHAIN
    ents[0]['size'] = len(mini)
    mf = b''.join(struct.pack('<I', x) for x in minifat)
    minifat_start = zincir(mf) if minifat else ENDOFCHAIN
    n_minifat = (len(mf) + SEC - 1) // SEC
    # dizin
    db = bytearray()
    for e in ents:
        ad = e['ad'].encode('utf-16le')
        assert len(ad) <= 62, e['ad']
        db += ad + b'\x00\x00' + b'\x00' * (64 - len(ad) - 2)
        db += struct.pack('<HBB', len(ad) + 2, e['tip'], 1)  # renk siyah
        db += struct.pack('<III', e['sol'], e['sag'], e['cocuk'])
        db += e['clsid'] + struct.pack('<I', 0) + b'\x00' * 16
        db += struct.pack('<II', e['start'] if e['start'] is not None else ENDOFCHAIN, e['size']) + b'\x00' * 4
    while len(db) % SEC:
        db += (b'\x00' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM)
               + b'\x00' * 36 + b'\x00' * 8 + b'\x00' * 4)[:128]
    dir_start = zincir(bytes(db))
    # FAT sektorleri
    nfat = 1
    while True:
        toplam = len(sectors) + nfat
        if nfat * (SEC // 4) >= toplam:
            break
        nfat += 1
    assert nfat <= 109, 'DIFAT gerekli'
    fat_start = len(sectors)
    for k in range(nfat):
        fat.append(FATSECT)
    fat += [FREESECT] * (nfat * (SEC // 4) - len(fat))
    fb = b''.join(struct.pack('<I', x) for x in fat)
    for k in range(nfat):
        sectors.append(fb[k * SEC:(k + 1) * SEC])
    difat = [fat_start + k for k in range(nfat)] + [FREESECT] * (109 - nfat)
    hdr = (b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\x00' * 16 + struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6)
           + b'\x00' * 6 + struct.pack('<IIIIIIIII', 0, nfat, dir_start, 0, CUTOFF, minifat_start, n_minifat,
                                       ENDOFCHAIN, 0)
           + b''.join(struct.pack('<I', x) for x in difat))
    assert len(hdr) == 512
    with open(yol, 'wb') as f:
        f.write(hdr)
        for s in sectors:
            f.write(s)


def oku(yol):
    """olefile ile mevcut dosyadan agac (CLSID'ler korunur)."""
    import olefile, uuid
    o = olefile.OleFileIO(yol)

    def dugum(de, yolu):
        cl = uuid.UUID(de.clsid).bytes_le if de.clsid else b'\x00' * 16
        n = {'ad': de.name, 'clsid': cl}
        if de.entry_type == olefile.STGTY_STREAM:
            n['veri'] = o.openstream(yolu).read()
        else:
            n['veri'] = None
            n['cocuk'] = [dugum(k, yolu + [k.name]) for k in de.kids]
        return n

    kok = {'ad': 'Root Entry', 'clsid': uuid.UUID(o.root.clsid).bytes_le if o.root.clsid else b'\x00' * 16, 'veri': None,
           'cocuk': [dugum(k, [k.name]) for k in o.root.kids]}
    o.close()
    return kok
