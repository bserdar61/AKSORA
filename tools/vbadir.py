# VBA 'dir' akisi kayitlari: [(id, veri)]. PROJECTVERSION (0x09) boyut alani 4 ama veri 6 bayttir.
import struct

def parse(d):
    recs = []
    i = 0
    while i + 6 <= len(d):
        rid, size = struct.unpack('<HI', d[i:i + 6])
        i += 6
        if rid == 0x09:
            size = 6
        recs.append([rid, d[i:i + size]])
        i += size
    return recs


def build(recs):
    out = bytearray()
    for rid, dt in recs:
        size = 4 if rid == 0x09 else len(dt)
        out += struct.pack('<HI', rid, size) + dt
    return bytes(out)


def modules(recs):
    """[(ad, akis_adi, offset_kayit_index)]"""
    res = []
    cur = None
    for k, (rid, dt) in enumerate(recs):
        if rid == 0x19:
            cur = {'ad': dt.decode('cp1254'), 'akis': None, 'ofs': None}
        elif rid == 0x1A and cur is not None:
            cur['akis'] = dt.decode('cp1254')
        elif rid == 0x31 and cur is not None:
            cur['ofs'] = k
        elif rid == 0x2B and cur is not None:
            res.append(cur); cur = None
    return res
