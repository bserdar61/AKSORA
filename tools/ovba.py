# MS-OVBA 2.4.1 sikistirma / acma.
import struct

def decompress(data):
    if not data or data[0] != 1:
        raise ValueError('ovba: imza yok')
    out = bytearray()
    pos = 1
    n = len(data)
    while pos < n:
        hdr = struct.unpack('<H', data[pos:pos + 2])[0]
        size = (hdr & 0x0FFF) + 3
        compressed = (hdr >> 15) & 1
        chunk_end = min(pos + size, n)
        pos += 2
        start = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]; pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos]); pos += 1
                else:
                    tok = struct.unpack('<H', data[pos:pos + 2])[0]; pos += 2
                    dlen = len(out) - start
                    bc = max((dlen - 1).bit_length(), 4)
                    lmask = 0xFFFF >> bc
                    off = (tok >> (16 - bc)) + 1
                    ln = (tok & lmask) + 3
                    s = len(out) - off
                    for i in range(ln):
                        out.append(out[s + i])
    return bytes(out)


def _chunk(src):
    # src: <= 4096 bayt. Donus: sikistirilmis parca (baslik dahil)
    out = bytearray()
    i = 0
    n = len(src)
    # basit sozluk: 3 baytlik onek -> konum listesi
    table = {}
    while i < n:
        flag_pos = len(out); out.append(0); flags = 0
        for bit in range(8):
            if i >= n:
                break
            bc = max((i - 1).bit_length(), 4) if i > 0 else 4
            lmask = 0xFFFF >> bc
            maxlen = lmask + 3
            best_len = 0; best_off = 0
            if i + 3 <= n:
                key = bytes(src[i:i + 3])
                cands = table.get(key)
                if cands:
                    maxoff = 1 << bc
                    for c in reversed(cands[-256:]):
                        off = i - c
                        if off > maxoff:
                            break
                        ln = 3
                        lim = min(maxlen, n - i)
                        while ln < lim and src[c + ln] == src[i + ln]:
                            ln += 1
                        if ln > best_len:
                            best_len = ln; best_off = off
                            if ln == lim:
                                break
            if best_len >= 3:
                tok = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack('<H', tok)
                flags |= 1 << bit
                for k in range(i, i + best_len):
                    if k + 3 <= n:
                        table.setdefault(bytes(src[k:k + 3]), []).append(k)
                i += best_len
            else:
                out.append(src[i])
                if i + 3 <= n:
                    table.setdefault(bytes(src[i:i + 3]), []).append(i)
                i += 1
        out[flag_pos] = flags
    if len(out) > 4096:
        # sikistirilmamis parca
        data = bytes(src) + b'\x00' * (4096 - n)
        return struct.pack('<H', 0x3000 | (4096 + 2 - 3)) + data
    hdr = 0xB000 | (len(out) + 2 - 3)
    return struct.pack('<H', hdr) + bytes(out)


def compress(data):
    out = bytearray(b'\x01')
    for p in range(0, len(data), 4096):
        out += _chunk(data[p:p + 4096])
    return bytes(out)


if __name__ == '__main__':
    import os, sys
    for t in [b'', b'a', b'abcabcabcabc' * 1000, os.urandom(10000), open(sys.argv[0], 'rb').read() * 3]:
        assert decompress(compress(t)) == t
    print('ovba ok')
