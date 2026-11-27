#!/usr/bin/env python3
# Derleme oncesi donusum: VBA metin sabitlerindeki ASCII disi karakterler ChrW$(kod) olur.
# Sebep: VBA kaynagi Windows'un "Unicode olmayan programlar" kod sayfasiyla okunur; Turkce olmayan
# PC'de "İ" -> "Ý", "Ş" -> "Þ" olur, sayfa/sekil adlari eslesmez (Hata 9). ChrW$ her dilde aynidir.
# Kullanim: tr_chrw.py <kaynak_klasor> <hedef_klasor>   (kaynak okunabilir kalir)
import re, sys, os, glob
STR = re.compile(r'"(?:[^"]|"")*"')
CONST = re.compile(r'^(\s*)Private Const (\w+) As String = ("(?:[^"]|"")*")\s*$')

def lit_expr(lit):
    body = lit[1:-1]
    parts, buf, i = [], '', 0
    while i < len(body):
        c = body[i]
        if ord(c) > 127:
            if buf: parts.append('"' + buf + '"'); buf = ''
            parts.append('ChrW$(%d)' % ord(c))
        else:
            buf += c
        i += 1
    if buf: parts.append('"' + buf + '"')
    return '(' + ' & '.join(parts) + ')'

def code_split(ln):
    # (kod, yorum) - yorum disindaki kisim
    inq = False
    for i, c in enumerate(ln):
        if c == '"': inq = not inq
        elif c == "'" and not inq: return ln[:i], ln[i:]
    return ln, ''

def bol(code, yorum):
    # ' & ' (metin disinda) noktalarindan ' _' ile bol; her parca <= 850
    ind = re.match(r'\s*', code).group(0) + '    '
    noktalar, inq = [], False
    for i, c in enumerate(code):
        if c == '"': inq = not inq
        elif not inq and code.startswith(' & ', i): noktalar.append(i + 2)
    parcalar, bas = [], 0
    while len(code) - bas > 850:
        aday = [k + 1 for k in noktalar if bas < k + 1 <= bas + 850]
        if not aday: break
        parcalar.append(code[bas:aday[-1]]); bas = aday[-1]
    parcalar.append(code[bas:])
    sonuc = []
    for j, pc in enumerate(parcalar):
        pc = pc if j == 0 else ind + pc.lstrip()
        sonuc.append(pc.rstrip() + ' _' if j < len(parcalar) - 1 else pc + yorum)
    return sonuc

def donustur(t):
    out, ekler, n = [], [], 0
    for ln in t.split('\r\n'):
        s = ln.lstrip()
        if s.upper().startswith('ATTRIBUTE ') or s.startswith("'") or s.upper().startswith('REM '):
            out.append(ln); continue
        code, yorum = code_split(ln)
        m = CONST.match(code)
        if m and any(ord(c) > 127 for c in m.group(3)):
            ekler.append('Private Function %s() As String\r\n    %s = %s\r\nEnd Function' % (m.group(2), m.group(2), lit_expr(m.group(3))))
            out.append(m.group(1) + "' (Const -> fonksiyon: " + m.group(2) + ")"); n += 1
            continue
        def rp(mm):
            nonlocal n
            l = mm.group(0)
            if any(ord(c) > 127 for c in l): n += 1; return lit_expr(l)
            return l
        code2 = STR.sub(rp, code)
        if 'Const ' in code2 and 'ChrW$' in code2:
            raise SystemExit('Desteklenmeyen Const: ' + ln)
        if len(code2 + yorum) > 900:
            out.extend(bol(code2, yorum))
        else:
            out.append(code2 + yorum)
    t2 = '\r\n'.join(out)
    devam = 0
    for ln in t2.split('\r\n'):
        c, _ = code_split(ln)
        if c.rstrip().endswith(' _'):
            devam += 1
            if devam > 24: raise SystemExit('24 satir devami asildi: ' + ln[:80])
        else:
            devam = 0
    if ekler:
        t2 = t2.rstrip('\r\n') + '\r\n\r\n' + '\r\n\r\n'.join(ekler) + '\r\n'
    return t2, n

if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    os.makedirs(dst, exist_ok=True)
    top, enuzun = 0, 0
    for p in sorted(glob.glob(os.path.join(src, '*.txt'))):
        t = open(p, encoding='cp1254', newline='').read()
        t2, n = donustur(t)
        top += n
        for ln in t2.split('\r\n'):
            if len(ln) > enuzun: enuzun = len(ln)
            if len(ln) > 1000: raise SystemExit('Satir cok uzun (%d): %s' % (len(ln), os.path.basename(p)))
        open(os.path.join(dst, os.path.basename(p)), 'w', encoding='cp1254', newline='').write(t2)
    print('donusen metin:', top, ' en uzun satir:', enuzun)
