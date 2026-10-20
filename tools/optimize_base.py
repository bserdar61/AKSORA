"""AKSORA base optimizer: packs il archive sheets (PAKET1), trims copy-bloat rows
on Y_*_STOK sheets. Usage: optimize_base.py in.xlsm out.xlsm"""
import sys, re, zipfile, html

src, out = sys.argv[1:3]
zi = zipfile.ZipFile(src)
files = {i.filename: zi.read(i.filename) for i in zi.infolist()}
infos = {i.filename: i for i in zi.infolist()}

wb = files['xl/workbook.xml'].decode('utf8')
rels = files['xl/_rels/workbook.xml.rels'].decode('utf8')


def sheet_path(rid):
    m = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    return 'xl/' + m.group(1).replace('/xl/', '').lstrip('/')


sheets = [(m.group(1), sheet_path(m.group(2))) for m in
          re.finditer(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb)]

# shared strings
sst = []
sx = files['xl/sharedStrings.xml'].decode('utf8')
for si in re.finditer(r'<si>(.*?)</si>', sx, re.S):
    parts = re.findall(r'<t(?: [^>]*)?>(.*?)</t>', si.group(1), re.S)
    sst.append(html.unescape(''.join(parts)))


def col_idx(ref):
    letters = re.match(r'[A-Z]+', ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def cell_value(attrs, body):
    if body is None:
        return None
    t = re.search(r' t="(\w+)"', attrs)
    t = t.group(1) if t else 'n'
    if t == 'inlineStr':
        parts = re.findall(r'<t(?: [^>]*)?>(.*?)</t>', body, re.S)
        return ('s', html.unescape(''.join(parts)))
    v = re.search(r'<v>(.*?)</v>', body, re.S)
    if not v:
        return None
    v = v.group(1)
    if t == 's':
        return ('s', sst[int(v)])
    if t in ('str',):
        return ('s', html.unescape(v))
    if t == 'b':
        return ('b', v)
    if t == 'e':
        return ('s', '#HATA')
    return ('n', v)


def field(val):
    if val is None:
        return ''
    k, v = val
    if k == 's':
        return 's' + v.replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
    if k == 'b':
        return 'b' + ('1' if v in ('1', 'true') else '0')
    return 'n' + v


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def pack_sheet(x):
    sd0 = x.index('<sheetData')
    sd1 = x.index('</sheetData>') + len('</sheetData>')
    body = x[sd0:sd1]
    rows = re.findall(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', body, re.S)
    row1 = ''
    data = []
    for r, content in rows:
        r = int(r)
        if r == 1:
            row1 = ''.join(m.group(0) for m in CELL.finditer(content or '') if m.group(1) in ('A', 'B'))
            continue
        vals = [None] * 14
        for m in CELL.finditer(content or ''):
            ci = col_idx(m.group(1))
            if ci <= 14:
                vals[ci - 1] = cell_value(m.group(3), m.group(4))
        if vals[4] is None:  # column E empty -> not a data row (matches VBA row count rule)
            continue
        data.append('\t'.join(field(v) for v in vals))
    # blocks <= 30000 chars, same rule as AKSORA_PAKET_YAZ
    blocks, cur, ln = [], [], 0
    for line in data:
        if ln + len(line) + 1 > 30000 and cur:
            blocks.append('\n'.join(cur)); cur, ln = [], 0
        cur.append(line); ln += len(line) + 1
    if cur:
        blocks.append('\n'.join(cur))
    out_rows = ['<row r="1">%s<c r="C1" t="inlineStr"><is><t>PAKET1</t></is></c><c r="D1"><v>%d</v></c></row>' % (row1, len(data))]
    for i, b in enumerate(blocks):
        r = i + 2
        out_rows.append('<row r="%d"><c r="A%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c></row>' % (r, r, esc(b)))
    newsd = '<sheetData>' + ''.join(out_rows) + '</sheetData>'
    x = x[:sd0] + newsd + x[sd1:]
    x = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:D%d"/>' % (len(blocks) + 1), x)
    return x, data, blocks


def paket2_lines(lines):
    """PAKET2: a field equal to the previous line's same field (non-empty) is written as '='."""
    out, prev = [], None
    for line in lines:
        f = line.split('\t')
        g = [('=' if i > 0 and prev is not None and i < len(prev) and x != '' and prev[i] == x else x) for i, x in enumerate(f)]
        out.append('\t'.join(g)); prev = f
    return out


def paket2_decode(lines):
    res, prev = [], [''] * 64
    for line in lines:
        f = line.split('\t')
        for i, x in enumerate(f):
            if x == '=':
                f[i] = prev[i]
            else:
                prev[i] = x
        res.append('\t'.join(f))
    return res


def paket2_sheet(x, name):
    sd0 = x.index('<sheetData'); sd1 = x.index('</sheetData>') + len('</sheetData>')
    body = x[sd0:sd1]
    rows = re.findall(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', body, re.S)
    row1 = ''; d1 = None; blocks = []
    for r, content in rows:
        r = int(r)
        cells = {m.group(1): m for m in CELL.finditer(content or '')}
        if r == 1:
            row1 = ''.join(m.group(0) for m in CELL.finditer(content or '') if m.group(1) in ('A', 'B'))
            if 'D' in cells:
                v = cell_value(cells['D'].group(3), cells['D'].group(4)); d1 = int(float(v[1]))
            continue
        if 'A' in cells:
            v = cell_value(cells['A'].group(3), cells['A'].group(4))
            if v and v[1] != '':
                blocks.append(v[1])
    lines = []
    for b in blocks:
        lines += b.split('\n')
    assert d1 is not None and len(lines) == d1, (name, d1, len(lines))
    new = paket2_lines(lines)
    assert paket2_decode(new) == lines, name
    nb, cur, ln = [], [], 0
    for line in new:
        if ln + len(line) + 1 > 30000 and cur:
            nb.append('\n'.join(cur)); cur, ln = [], 0
        cur.append(line); ln += len(line) + 1
    if cur:
        nb.append('\n'.join(cur))
    back = []
    for b in nb:
        back += b.split('\n')
    assert paket2_decode(back) == lines, name
    out_rows = ['<row r="1">%s<c r="C1" t="inlineStr"><is><t>PAKET2</t></is></c><c r="D1"><v>%d</v></c></row>' % (row1, d1)]
    for i, b in enumerate(nb):
        r = i + 2
        out_rows.append('<row r="%d"><c r="A%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c></row>' % (r, r, esc(b)))
    x2 = x[:sd0] + '<sheetData>' + ''.join(out_rows) + '</sheetData>' + x[sd1:]
    x2 = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:D%d"/>' % (len(nb) + 1), x2)
    return x2, sum(len(b) for b in blocks), sum(len(b) for b in nb)


def unpack(blocks):
    rows = []
    for b in blocks:
        for line in b.split('\n'):
            rows.append(line.split('\t'))
    return rows


def trim_rows(x, keep_min=60):
    sd0 = x.index('<sheetData')
    sd1 = x.index('</sheetData>')
    body = x[sd0:sd1]
    last = 0
    for m in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', body, re.S):
        r = int(m.group(1)); content = m.group(2) or ''
        for c in CELL.finditer(content):
            if col_idx(c.group(1)) <= 14 and c.group(4) and re.search(r'<(v|is)\b', c.group(4)):
                last = max(last, r)
    keep = max(last + 2, keep_min)
    removed = [0]

    def sub(m):
        if int(m.group(1)) > keep:
            removed[0] += 1
            return ''
        return m.group(0)
    body2 = re.sub(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', sub, body, flags=re.S)
    x = x[:sd0] + body2 + x[sd1:]
    x = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"/>', lambda m: '<dimension ref="%s:%s%d"/>' % (m.group(1), m.group(2), keep), x)
    return x, keep, removed[0]


report = []
for name, p in sheets:
    x = files[p].decode('utf8')
    m1 = re.search(r'<c r="C1"([^>]*?)(?:/>|>(.*?)</c>)', x, re.S)
    v1 = cell_value(m1.group(1), m1.group(2)) if m1 else None
    paketli = bool(v1) and str(v1[1]).startswith('PAKET')
    if (name.startswith('_MA_') or name.startswith('_YA_') or name.startswith('_IA_')) and paketli and v1[1] == 'PAKET1':
        nx, c1, c2 = paket2_sheet(x, name)
        files[p] = nx.encode('utf8')
        report.append('%s: PAKET2 %d -> %d karakter' % (name, c1, c2))
    elif (name.startswith('_MA_') or name.startswith('_YA_')) and paketli:
        report.append('%s: zaten paketli, dokunulmadi' % name)
    elif name.startswith('_MA_') or name.startswith('_YA_'):
        nx, data, blocks = pack_sheet(x)
        # verification: unpacked rows equal packed lines
        back = unpack(blocks) if blocks else []
        assert len(back) == len(data), name
        for a, b in zip(back, data):
            assert '\t'.join(a) == b
        files[p] = nx.encode('utf8')
        report.append('%s: %d satir, %d blok, %d KB -> %d KB' % (name, len(data), len(blocks), len(x) // 1024, len(nx) // 1024))
    elif (name.startswith('Y_') or name.startswith('I_')) and name.endswith('_STOK'):
        nx, keep, rem = trim_rows(x)
        files[p] = nx.encode('utf8')
        report.append('%s: %d satir silindi (son %d), %d KB -> %d KB' % (name, rem, keep, len(x) // 1024, len(nx) // 1024))

# --- kullanilmayan paylasimli metinler atilir (dosya ilk halinde de kucuk olur)
SCELL = re.compile(r'(<c r="[A-Z]+\d+"[^>]*? t="s"[^>]*>(?:<f>[^<]*</f>|<f [^>]*/>|<f [^>]*>[^<]*</f>)?<v>)(\d+)(</v>)')
si_raw = [m.group(0) for m in re.finditer(r'<si>.*?</si>', sx, re.S)]
assert len(si_raw) == len(sst)
ws_paths = [p2 for p2 in files if p2.startswith('xl/worksheets/sheet') and p2.endswith('.xml')]
used = set()
for p2 in ws_paths:
    for m in SCELL.finditer(files[p2].decode('utf8')):
        used.add(int(m.group(2)))
alls = set()
for p2 in ws_paths:
    alls.update(int(v) for v in re.findall(r' t="s"[^>]*><v>(\d+)</v>', files[p2].decode('utf8')))
assert alls <= used, 'paylasimli metin deseni eksik'
order = sorted(used)
remap = {o: n for n, o in enumerate(order)}
newsst = [sst[o] for o in order]
for p2 in ws_paths:
    x2 = files[p2].decode('utf8')
    before = [sst[int(m.group(2))] for m in SCELL.finditer(x2)]
    x3 = SCELL.sub(lambda m: m.group(1) + str(remap[int(m.group(2))]) + m.group(3), x2)
    after = [newsst[int(m.group(2))] for m in SCELL.finditer(x3)]
    assert before == after, p2
    files[p2] = x3.encode('utf8')
refs = sum(len(SCELL.findall(files[p2].decode('utf8'))) for p2 in ws_paths)
head = re.match(r'(.*?<sst[^>]*>)', sx, re.S).group(1)
head = re.sub(r'count="\d+"', 'count="%d"' % refs, head)
head = re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % len(used), head)
files['xl/sharedStrings.xml'] = (head + ''.join(si_raw[o] for o in order) + '</sst>').encode('utf8')
report.append('paylasimli metin: %d -> %d' % (len(si_raw), len(used)))

zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
for fn, data in files.items():
    zo.writestr(infos[fn], data, compress_type=zipfile.ZIP_DEFLATED)
zo.close()
print('\n'.join(report))


# ---------------- images ----------------
def optimize_images(path_in, path_out):
    import io, posixpath
    from PIL import Image
    z = zipfile.ZipFile(path_in)
    fl = {i.filename: z.read(i.filename) for i in z.infolist()}
    inf = {i.filename: i for i in z.infolist()}
    # media -> list of (needed_px_w, needed_px_h) ; and referencing rels files
    need, refs = {}, {}
    for fn in fl:
        if fn.endswith('.rels'):
            txt = fl[fn].decode('utf8')
            for m in re.finditer(r'Target="([^"]*media/[^"]+)"', txt):
                base = posixpath.dirname(posixpath.dirname(fn))
                tgt = posixpath.normpath(posixpath.join(base, m.group(1))) if not m.group(1).startswith('/') else m.group(1).lstrip('/')
                refs.setdefault(tgt, set()).add(fn)
    for fn in fl:
        if re.match(r'xl/drawings/drawing\d+\.xml$', fn):
            x = fl[fn].decode('utf8')
            rp = 'xl/drawings/_rels/' + posixpath.basename(fn) + '.rels'
            rm = {}
            if rp in fl:
                for m in re.finditer(r'<Relationship [^>]*?Id="(rId\d+)"[^>]*?Target="([^"]+)"', fl[rp].decode('utf8')):
                    rm[m.group(1)] = posixpath.normpath(posixpath.join('xl/drawings', m.group(2)))
                for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="(rId\d+)"', fl[rp].decode('utf8')):
                    rm[m.group(2)] = posixpath.normpath(posixpath.join('xl/drawings', m.group(1)))
            for pm in re.finditer(r'<xdr:pic>.*?</xdr:pic>', x, re.S):
                b = pm.group(0)
                e = re.search(r'r:embed="(rId\d+)"', b)
                ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', b)
                if not e or not ext or e.group(1) not in rm:
                    continue
                sr = re.search(r'<a:srcRect([^/]*)/>', b)
                fx = fy = 1.0
                if sr:
                    g = dict(re.findall(r'(\w)="(-?\d+)"', sr.group(1)))
                    fx = max(0.05, 1 - (int(g.get('l', 0)) + int(g.get('r', 0))) / 100000)
                    fy = max(0.05, 1 - (int(g.get('t', 0)) + int(g.get('b', 0))) / 100000)
                w = int(ext.group(1)) / 9525 * 1.5 / fx
                h = int(ext.group(2)) / 9525 * 1.5 / fy
                k = rm[e.group(1)]
                ow, oh = need.get(k, (0, 0))
                need[k] = (max(ow, w), max(oh, h))
    rename = {}
    rep = []
    for media, (nw, nh) in need.items():
        if media not in fl:
            continue
        # only drawing references allowed for renaming
        if any(not r.startswith('xl/drawings/_rels/') for r in refs.get(media, ())):
            continue
        data = fl[media]
        try:
            im = Image.open(io.BytesIO(data)); im.load()
        except Exception:
            continue
        w0, h0 = im.size
        sc = min(1.0, max(nw / w0, nh / h0) * 1.0)
        img = im
        if sc < 0.87:
            img = im.resize((max(1, int(w0 * sc)), max(1, int(h0 * sc))), Image.LANCZOS)
        has_alpha = img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info)
        if has_alpha and img.mode == 'RGBA':
            if img.getchannel('A').getextrema()[0] == 255:
                has_alpha = False
        bio = io.BytesIO()
        if not has_alpha and len(data) > 150000:
            img.convert('RGB').save(bio, 'JPEG', quality=90, optimize=True, progressive=False)
            newname = re.sub(r'\.\w+$', '.jpeg', media)
            if newname != media and newname in fl:
                newname = re.sub(r'\.\w+$', '_o.jpeg', media)
        else:
            if img is im and media.lower().endswith('.png'):
                img.save(bio, 'PNG', optimize=True)
            elif media.lower().endswith('.png'):
                img.save(bio, 'PNG', optimize=True)
            else:
                img.convert('RGB').save(bio, 'JPEG', quality=90, optimize=True)
            newname = media
        nd = bio.getvalue()
        if len(nd) >= len(data) * 0.95:
            continue
        del fl[media]
        fl[newname] = nd
        if newname != media:
            rename[media] = newname
            inf[newname] = zipfile.ZipInfo(newname, date_time=inf[media].date_time)
        rep.append('%s %dx%d %dKB -> %s %dx%d %dKB' % (media, w0, h0, len(data) // 1024, newname, img.size[0], img.size[1], len(nd) // 1024))
    for old, new in rename.items():
        ob, nb = posixpath.basename(old), posixpath.basename(new)
        for r in refs[old]:
            t = fl[r].decode('utf8').replace('/' + ob + '"', '/' + nb + '"')
            fl[r] = t.encode('utf8')
    ct = fl['[Content_Types].xml'].decode('utf8')
    if 'Extension="jpeg"' not in ct:
        ct = ct.replace('<Default Extension="xml"', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="xml"')
        fl['[Content_Types].xml'] = ct.encode('utf8')
    zo = zipfile.ZipFile(path_out, 'w', zipfile.ZIP_DEFLATED)
    order = [i.filename for i in z.infolist() if i.filename in fl] + [n for n in fl if n not in inf or n in rename.values()]
    seen = set()
    for fn in order:
        if fn in seen:
            continue
        seen.add(fn)
        zo.writestr(inf.get(fn, zipfile.ZipInfo(fn)), fl[fn], compress_type=zipfile.ZIP_DEFLATED)
    zo.close()
    print('\n'.join(rep))


if len(sys.argv) > 3 and sys.argv[3] == '--img':
    optimize_images(out, out + '.img.xlsm')
