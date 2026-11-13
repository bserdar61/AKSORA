#!/usr/bin/env python3
"""AKSORA v2 görünüm prototipi (gerçek Excel dosyası, makrosuz, örnek veriyle).

Amaç: tasarımı HTML çizimiyle değil, Excel'in kendi öğeleriyle göstermek.
Dosyadaki her şey hücre, şekil (metin kutusu), grafik ve köprüden ibarettir;
yani VBA'nın gerçek AKSORA sayfalarında birebir çizebileceği öğelerdir.

Kullanım:  python3 v2/prototip_uret.py v2/AKSORA_v2_Prototip.xlsx
Gereken :  pip install xlsxwriter
"""
import re
import shutil
import sys
import zipfile

import xlsxwriter

# ---------------------------------------------------------------- renkler
NAVY, SIDE, SIDE_ON = "#0B2541", "#0A1F38", "#12385F"
TEAL, TEAL_D = "#00A38D", "#007D6E"
INK, INK2, INK3 = "#0F172A", "#475569", "#94A3B8"
LINE, BG, WHITE, HEAD = "#E5EAF0", "#F3F6FA", "#FFFFFF", "#F8FAFC"
RED, AMB, GRN, BLU, VIO = "#DC2626", "#D97706", "#16A34A", "#2563EB", "#8B5CF6"
REDBG, AMBBG, GRNBG, BLUBG, SLBG = "#FEE2E2", "#FEF3C7", "#DCFCE7", "#DBEAFE", "#EEF2F6"
FONT, ICON = "Segoe UI", "Segoe MDL2 Assets"

# Windows simge yazı tipi (Segoe MDL2 Assets) karakterleri
IC = dict(home="", search="", cart="", swap="", clock="", box="",
          pin="", people="", chart="", sync="", bell="", shield="",
          doc="", down="", right="", back="", settings="", list="",
          compare="")

SW, TOP, X0, CW = 220, 64, 244, 1240          # kenar menü, üst bant, içerik sol kenarı ve genişliği (piksel)


class Sayfa:
    """Bir çalışma sayfası üzerinde piksel konumlu çizim yardımcıları."""

    def __init__(self, wb, ad, fm):
        self.wb, self.ws, self.fm = wb, wb.add_worksheet(ad), fm
        self.ws.hide_gridlines(2)
        self.ws.hide_row_col_headers()
        self.ws.set_zoom(100)
        self.ws.set_landscape()
        self.ws.set_paper(8)          # A3: çıktı tek sayfaya sığar
        self.ws.fit_to_pages(1, 1)

    # -- şekil: etiket = köşe yarıçapı / gölge / iç boşluk bilgisi (kayıt sonrası XML'e işlenir)
    def kutu(self, x, y, w, h, metin="", dolgu=None, cizgi=None, renk=INK, boyut=10, kalin=False, yatay="left",
             dikey="middle", font=FONT, url=None, r=0, golge=False, ic=8, ipucu=None):
        o = {"x_offset": x, "y_offset": y, "width": w, "height": h, "object_position": 3,
             "fill": {"color": dolgu} if dolgu else {"none": True},
             "line": {"color": cizgi, "width": 0.75} if cizgi else {"none": True},
             "font": {"name": font, "size": boyut, "bold": kalin, "color": renk},
             "align": {"vertical": dikey, "horizontal": yatay},
             "description": f"r={r};g={int(golge)};i={ic}"}
        if url:
            o["url"] = url
        if ipucu:
            o["tip"] = ipucu
        self.ws.insert_textbox(0, 0, metin, o)

    def simge(self, x, y, ad, renk=INK2, boyut=12, w=26, h=26, dolgu=None, r=0, url=None):
        self.kutu(x, y, w, h, IC[ad], dolgu=dolgu, renk=renk, boyut=boyut, yatay="center", font=ICON, r=r, ic=0, url=url)

    def kart(self, x, y, w, h):
        self.kutu(x, y, w, h, dolgu=WHITE, cizgi=LINE, r=10, golge=True)

    def grafik(self, ch, x, y, w, h):
        ch.set_size({"width": w, "height": h})
        self.ws.insert_chart(0, 0, ch, {"x_offset": x, "y_offset": y, "object_position": 3})


def bicimler(wb):
    f = {}
    def ekle(ad, **k):
        k.setdefault("font_name", FONT)
        k.setdefault("font_size", 10)
        k.setdefault("valign", "vcenter")
        f[ad] = wb.add_format(k)
    ekle("bg", bg_color=BG)
    ekle("side", bg_color=SIDE)
    ekle("top", bg_color=WHITE)
    ekle("topb", bg_color=WHITE, bottom=1, bottom_color=LINE)
    ekle("navy", bg_color=NAVY)
    ekle("th", bg_color=HEAD, font_color=INK2, bold=True, font_size=9, bottom=1, bottom_color=LINE, top=1, top_color=LINE, indent=1)
    ekle("thr", bg_color=HEAD, font_color=INK2, bold=True, font_size=9, bottom=1, bottom_color=LINE, top=1, top_color=LINE, align="right", indent=1)
    ekle("td", bg_color=WHITE, font_color=INK, bottom=1, bottom_color="#F0F3F7", indent=1)
    ekle("tdb", bg_color=WHITE, font_color=INK, bold=True, bottom=1, bottom_color="#F0F3F7", indent=1)
    ekle("tdm", bg_color=WHITE, font_color=INK3, bottom=1, bottom_color="#F0F3F7", indent=1)
    ekle("tdn", bg_color=WHITE, font_color=INK, bottom=1, bottom_color="#F0F3F7", align="right", num_format="#,##0.00", indent=1)
    ekle("tdi", bg_color=WHITE, font_color=INK, bottom=1, bottom_color="#F0F3F7", align="right", num_format="#,##0", indent=1)
    ekle("tdp", bg_color=WHITE, font_color=INK, bottom=1, bottom_color="#F0F3F7", align="right", num_format="0.0%", indent=1)
    for ad, renk in (("r", RED), ("a", AMB), ("g", GRN), ("b", BLU), ("s", INK2)):
        ekle("st" + ad, bg_color=WHITE, font_color=renk, bold=True, bottom=1, bottom_color="#F0F3F7", indent=1)
    ekle("ak", bg_color=WHITE)
    ekle("etiket", bg_color=WHITE, font_color=INK2, bold=True, font_size=9, indent=1)
    ekle("tdq", bg_color=WHITE, font_color=INK, bottom=1, bottom_color="#F0F3F7", align="right", indent=1)
    ekle("giris", bg_color=WHITE, font_color=INK, font_size=13, bold=True, border=1, border_color="#CBD5E1", indent=1)
    return f


# ---------------------------------------------------------------- kabuk (kenar menü + üst bant)
MENU = [("GENEL", [("home", "Ana panel", "Ana Panel", None), ("search", "Malzeme arama", "Malzeme Arama", None)]),
        ("KARAR", [("cart", "Satın alma", "Satın Alma", ("37", RED)), ("swap", "Transfer listesi", None, ("286", None)),
                   ("clock", "Atıl stok", "Atıl Stok", ("1.506", None)), ("compare", "Stok karşılaştır", None, None)]),
        ("STOKLAR", [("box", "Tüm stoklar", "Stoklar", None), ("list", "Depolar", None, None),
                     ("pin", "İlçeler", None, None), ("people", "Yükleniciler", None, None)]),
        ("RAPOR", [("chart", "Raporlar", "Raporlar", None), ("shield", "Yönetim merkezi", "Yönetim Merkezi", None)])]


def kabuk(s, etkin, baslik, yol, ana_eylem=None, ikincil=None, satir=60, sutunlar=None):
    ws, f = s.ws, s.fm
    ws.set_column_pixels(0, 0, SW, f["side"])
    ws.set_column_pixels(1, 1, 24, f["bg"])
    # Sütun genişlikleri şekillerden ÖNCE ayarlanır: şekil konumları o anki genişliklere göre hesaplanır.
    sutunlar = sutunlar or []
    for j, g in enumerate(sutunlar):
        ws.set_column_pixels(2 + j, 2 + j, g, f["bg"])
    ws.set_column_pixels(2 + len(sutunlar), 60, 64, f["bg"])
    for r, h in enumerate((22, 22, 20)):
        ws.set_row_pixels(r, h, f["top"] if r < 2 else f["topb"])
        ws.write_blank(r, 0, None, f["side"])
    for r in range(3, satir):
        ws.set_row_pixels(r, 30, f["bg"])
        ws.write_blank(r, 0, None, f["side"])
    ws.freeze_panes(3, 1)
    # --- kenar menü
    s.kutu(18, 16, 34, 34, "A", dolgu=TEAL, renk=WHITE, boyut=15, kalin=True, yatay="center", r=9, ic=0)
    s.kutu(60, 12, 150, 24, "AKSORA", renk=WHITE, boyut=13, kalin=True, ic=0)
    s.kutu(60, 32, 150, 18, "Malzeme Stok Birimi", renk="#8FA3BC", boyut=8, ic=0)
    y = 70
    for grup, ogeler in MENU:
        s.kutu(18, y, 180, 20, grup, renk="#7C92AD", boyut=7.5, kalin=True, ic=0)
        y += 24
        for sim, ad, hedef, rozet in ogeler:
            ac = ad == etkin
            url = f"internal:'{hedef}'!A1" if hedef else None
            s.kutu(10, y, 200, 32, "          " + ad, dolgu=SIDE_ON if ac else SIDE, renk=WHITE if ac else "#C7D3E2",
                   boyut=9.5, kalin=ac, r=7, url=url, ic=4)
            if ac:
                s.kutu(10, y + 4, 3, 24, dolgu="#19C2A8", ic=0)
            s.simge(18, y + 3, sim, renk=WHITE if ac else "#9FB3CB", boyut=11, url=url)
            if rozet:
                s.kutu(162, y + 7, 40, 18, rozet[0], dolgu=rozet[1] or "#1C3F66", renk=WHITE, boyut=7.5, kalin=True,
                       yatay="center", r=9, ic=0)
            y += 34
        y += 6
    s.kutu(10, y + 6, 200, 32, "          Birim seçimi", dolgu=SIDE, renk="#9FB3CB", boyut=9.5, r=7, url="internal:'Giriş'!A1", ic=4)
    s.simge(18, y + 9, "back", renk="#9FB3CB", boyut=11)
    # --- üst bant
    s.kutu(X0, 6, 500, 18, yol, renk=INK3, boyut=8.5, ic=0)
    s.kutu(X0, 24, 500, 30, baslik, renk=INK, boyut=15, kalin=True, ic=0)
    x = X0 + CW
    s.simge(x - 34, 16, "bell", renk=INK2, boyut=11, w=34, h=32, dolgu=WHITE, r=8)
    x -= 44
    if ana_eylem:
        w = 60 + 7 * len(ana_eylem)
        x -= w
        s.kutu(x, 16, w, 32, "       " + ana_eylem, dolgu=TEAL_D, renk=WHITE, boyut=9.5, kalin=True, r=8, golge=True, ic=6)
        s.simge(x + 6, 19, "sync", renk=WHITE, boyut=10)
        x -= 10
    for ad in (ikincil or []):
        w = 30 + 7 * len(ad)
        x -= w
        s.kutu(x, 16, w, 32, ad, dolgu=WHITE, cizgi=LINE, renk=INK, boyut=9.5, kalin=True, yatay="center", r=8)
        x -= 10
    x -= 110
    s.kutu(x, 16, 110, 32, "       Artvin", dolgu=WHITE, cizgi=LINE, renk=INK, boyut=9.5, kalin=True, r=8, ic=6)
    s.simge(x + 4, 19, "pin", renk=INK2, boyut=10)
    s.simge(x + 82, 19, "down", renk=INK3, boyut=8)
    x -= 230
    s.kutu(x, 16, 220, 32, "       Malzeme kodu ya da adı", dolgu=BG, cizgi=LINE, renk=INK3, boyut=9.5, r=8, ic=6,
           url="internal:'Malzeme Arama'!A1")
    s.simge(x + 4, 19, "search", renk=INK3, boyut=10)


def mini_cizgi(wb, veri_ws, satir, n, renk):
    ch = wb.add_chart({"type": "line"})
    ch.add_series({"values": ["_veri", satir, 1, satir, n], "line": {"color": renk, "width": 2.25}, "marker": {"type": "none"}})
    ch.set_legend({"none": True})
    ch.set_x_axis({"visible": False})
    ch.set_y_axis({"visible": False, "major_gridlines": {"visible": False}})
    ch.set_chartarea({"border": {"none": True}, "fill": {"none": True}})
    ch.set_plotarea({"border": {"none": True}, "fill": {"none": True}})
    return ch


def kpi_satiri(s, y, ogeler, veri_satir=None):
    """ogeler: (simge, simge zemini, simge rengi, etiket, değer, rozet metni, rozet (zemin, yazı), alt metin, eğilim satırı)"""
    n = len(ogeler)
    w = (CW - 14 * (n - 1)) // n
    for i, (sim, zemin, renk, etiket, deger, rozet, rz, alt, egilim) in enumerate(ogeler):
        x = X0 + i * (w + 14)
        s.kart(x, y, w, 104)
        s.simge(x + 14, y + 12, sim, renk=renk, boyut=10, w=28, h=28, dolgu=zemin, r=8)
        s.kutu(x + 50, y + 12, w - 60, 28, etiket, renk=INK2, boyut=9, kalin=True, ic=0)
        s.kutu(x + 14, y + 44, w - 110, 34, deger, renk=INK, boyut=17, kalin=True, ic=0)
        cx = x + 14
        if rozet:
            rw = 16 + 6 * len(rozet)
            s.kutu(cx, y + 78, rw, 18, rozet, dolgu=rz[0], renk=rz[1], boyut=7.5, kalin=True, yatay="center", r=8, ic=0)
            cx += rw + 6
        if alt:
            s.kutu(cx, y + 78, 140, 18, alt, renk=INK3, boyut=8, ic=0)
        if egilim is not None:
            s.grafik(mini_cizgi(s.wb, None, egilim, 7, renk), x + w - 104, y + 40, 96, 56)
    return y + 104


def tablo(s, ust_satir, basliklar, satirlar, genislikler, bicim, ilk_sutun=2, suz=False):
    ws, f = s.ws, s.fm
    ws.set_row_pixels(ust_satir, 30)
    for j, b in enumerate(basliklar):
        ws.write(ust_satir, ilk_sutun + j, b.upper().replace("İ", "İ"), f["thr" if bicim[j] in ("tdn", "tdi", "tdp", "tdq") else "th"])
    for i, satir in enumerate(satirlar):
        ws.set_row_pixels(ust_satir + 1 + i, 30)
        for j, v in enumerate(satir):
            b = bicim[j]
            if b == "st":
                metin, tur = v
                ws.write(ust_satir + 1 + i, ilk_sutun + j, "●  " + metin, f["st" + tur])
            else:
                ws.write(ust_satir + 1 + i, ilk_sutun + j, v, f[b])
    if suz:
        ws.autofilter(ust_satir, ilk_sutun, ust_satir + len(satirlar), ilk_sutun + len(basliklar) - 1)
    return ust_satir + 1 + len(satirlar)


def sekmeler(s, x, y, ogeler, etkin=0):
    w = sum(24 + 7 * len(o) for o in ogeler) + 8
    s.kutu(x, y, w, 34, dolgu=SLBG, r=9)
    cx = x + 4
    for i, o in enumerate(ogeler):
        ow = 24 + 7 * len(o)
        if i == etkin:
            s.kutu(cx, y + 4, ow, 26, o, dolgu=WHITE, renk=INK, boyut=9, kalin=True, yatay="center", r=7, golge=True, ic=0)
        else:
            s.kutu(cx, y + 4, ow, 26, o, renk=INK2, boyut=9, kalin=True, yatay="center", ic=0)
        cx += ow
    return x + w


def y_satir(y):
    """Piksel y konumuna denk gelen satır (üst bant 64 px, sonrası 30 px)."""
    return 3 + (y - TOP + 29) // 30


# ---------------------------------------------------------------- sayfalar
def veri_sayfasi(wb):
    v = wb.add_worksheet("_veri")
    seriler = [[176, 175, 175.5, 174, 173.6, 173, 172.5], [48, 46, 45, 44, 41, 41, 37], [2.9, 2.95, 3.0, 3.0, 3.05, 3.06, 3.1],
               [0.9, 1.0, 1.05, 1.1, 1.2, 1.18, 1.16], [36, 35, 34.4, 33.6, 33, 32.4, 31.9], [2.4, 2.5, 2.6, 2.7, 2.8, 2.85, 2.9],
               [1.1, 1.2, 1.25, 1.3, 1.35, 1.4, 1.44], [20, 20.4, 20.6, 20.8, 21, 21, 21.1]]
    for i, sr in enumerate(seriler):
        v.write_row(i, 1, sr)
    v.write_column(10, 0, ["Yüklenici", "EPP / 7055", "CDP / 7155", "İlçeler"])
    v.write_column(10, 1, [93.5, 52.7, 20.0, 6.3])
    v.write_column(16, 0, ["İşletme Ambarı", "Müş.Bağ.Ambarı", "Sayaç Dm Ambar", "İSG-Hırdavat", "Planlı Bakım", "Diğer"])
    v.write_column(16, 1, [8.08, 11.49, 0.29, 0.09, 0.04, 0.02])
    v.write_column(24, 0, ["Artvin", "Trabzon", "Gümüşhane", "Giresun", "Rize"])
    v.write_column(24, 1, [172.5, 498.2, 141.9, 318.4, 258.1])
    v.write_column(24, 2, [1.8, 2.0, 3.2, 2.8, 2.1])
    gunler = [f"{d:02d}.{m:02d}" for m, d in [(10, x) for x in range(15, 32, 2)] + [(11, x) for x in range(2, 14, 2)]]
    v.write_column(32, 0, gunler)
    v.write_column(32, 1, [1410, 1409, 1407, 1406, 1404, 1401, 1400, 1398, 1397, 1395, 1393, 1392, 1391, 1390, 1389][:len(gunler)])
    v.write_column(32, 2, [52, 53, 55, 54, 56, 58, 57, 59, 61, 60, 62, 63, 62, 64, 64][:len(gunler)])
    v.hide()
    return len(gunler)


def pasta(wb, ilk, n, renkler, delik=62):
    ch = wb.add_chart({"type": "doughnut"})
    ch.add_series({"categories": ["_veri", ilk, 0, ilk + n - 1, 0], "values": ["_veri", ilk, 1, ilk + n - 1, 1],
                   "points": [{"fill": {"color": c}, "border": {"color": WHITE, "width": 1.5}} for c in renkler]})
    ch.set_hole_size(delik)
    ch.set_legend({"position": "right", "font": {"name": FONT, "size": 9, "color": INK2}})
    ch.set_chartarea({"border": {"none": True}, "fill": {"none": True}})
    return ch


def sutun(wb, ilk, n, sutun_no, renk, bicim="0", tur="column"):
    ch = wb.add_chart({"type": tur})
    ch.add_series({"categories": ["_veri", ilk, 0, ilk + n - 1, 0], "values": ["_veri", ilk, sutun_no, ilk + n - 1, sutun_no],
                   "fill": {"color": renk}, "border": {"none": True}, "gap": 80,
                   "data_labels": {"value": True, "font": {"name": FONT, "size": 8, "color": INK2}, "num_format": bicim}})
    ch.set_legend({"none": True})
    eksen = {"num_font": {"name": FONT, "size": 8, "color": INK3}, "line": {"none": True}, "major_tick_mark": "none"}
    ch.set_x_axis(dict(eksen))
    ch.set_y_axis(dict(eksen, visible=False, major_gridlines={"visible": True, "line": {"color": "#EEF2F6", "width": 0.75}}))
    ch.set_chartarea({"border": {"none": True}, "fill": {"none": True}})
    ch.set_plotarea({"border": {"none": True}, "fill": {"none": True}})
    return ch


def cizgi(wb, ilk, n, sutun_no, renk):
    ch = wb.add_chart({"type": "line"})
    ch.add_series({"categories": ["_veri", ilk, 0, ilk + n - 1, 0], "values": ["_veri", ilk, sutun_no, ilk + n - 1, sutun_no],
                   "line": {"color": renk, "width": 2.25}, "marker": {"type": "none"}})
    ch.set_legend({"none": True})
    eksen = {"num_font": {"name": FONT, "size": 8, "color": INK3}, "line": {"none": True}, "major_tick_mark": "none"}
    ch.set_x_axis(dict(eksen, interval_unit=3))
    ch.set_y_axis(dict(eksen, major_gridlines={"visible": True, "line": {"color": "#EEF2F6", "width": 0.75}}))
    ch.set_chartarea({"border": {"none": True}, "fill": {"none": True}})
    ch.set_plotarea({"border": {"none": True}, "fill": {"none": True}})
    return ch


def giris(wb, f):
    s = Sayfa(wb, "Giriş", f)
    ws = s.ws
    ws.set_column_pixels(0, 30, 64, f["navy"])
    for r in range(0, 40):
        ws.set_row_pixels(r, 24, f["navy"])
    s.kutu(60, 40, 40, 40, "A", dolgu=TEAL, renk=WHITE, boyut=17, kalin=True, yatay="center", r=10, ic=0)
    s.kutu(110, 36, 300, 28, "AKSORA", renk=WHITE, boyut=15, kalin=True, ic=0)
    s.kutu(110, 60, 300, 20, "Çoruh Elektrik Dağıtım A.Ş.", renk="#9FB3CB", boyut=8.5, ic=0)
    s.kutu(1260, 44, 170, 34, "       Yönetim merkezi", cizgi="#3B5878", renk=WHITE, boyut=9.5, kalin=True, r=8, ic=6,
           url="internal:'Yönetim Merkezi'!A1")
    s.simge(1266, 48, "shield", renk=WHITE, boyut=10)
    s.kutu(60, 150, 600, 44, "Hoş geldiniz", renk=WHITE, boyut=22, kalin=True, ic=0)
    s.kutu(60, 196, 900, 24, "Hangi birim için çalışacaksınız?   •   Son güncelleme 13.11.2026 08:44   •   5 il güncel",
           renk="#9FB3CB", boyut=10, ic=0)
    birimler = [("box", "Malzeme Stok", "Stok takibi, satın alma, transfer ve atıl stok kararları", "3 iş bekliyor", True),
                ("list", "Yatırım", "Hakediş, teminat ve yüklenici stok takibi", "1 teminat uyarısı", False),
                ("people", "Müşteri", "Seri numaralı malzeme, ilçe ve yüklenici kontrolleri", "Güncel", False),
                ("pin", "İlçe Yönetim", "İlçe depoları, demontaj kontrolü ve İSG zimmet", "834 zimmet bekliyor", False)]
    for i, (sim, ad, acik, durum, sec) in enumerate(birimler):
        x = 60 + i * 346
        url = "internal:'Ana Panel'!A1"
        s.kutu(x, 250, 330, 190, dolgu=WHITE if sec else "#14365C", cizgi=None if sec else "#2A4A70", r=12, golge=sec, url=url)
        s.simge(x + 20, 270, sim, renk=TEAL_D if sec else "#5EEAD4", boyut=14, w=42, h=42, dolgu="#E6FAF6" if sec else "#1B4470", r=11)
        s.kutu(x + 20, 322, 290, 28, ad, renk=INK if sec else WHITE, boyut=12.5, kalin=True, ic=0)
        s.kutu(x + 20, 350, 290, 40, acik, renk=INK2 if sec else "#9FB3CB", boyut=9, dikey="top", ic=0)
        s.kutu(x + 20, 398, 180, 26, durum, renk=INK3 if sec else "#7C92AD", boyut=8.5, ic=0)
        s.kutu(x + 240, 396, 70, 30, "Aç  →", dolgu=TEAL_D if sec else None, renk=WHITE if sec else "#5EEAD4", boyut=9.5,
               kalin=True, yatay="center", r=7, url=url, ic=0)
    s.kutu(60, 520, 600, 22, "Artvin  ·  Rize  ·  Trabzon  ·  Gümüşhane  ·  Giresun", renk="#7C92AD", boyut=8.5, ic=0)
    s.kutu(1330, 520, 100, 22, "v2 prototip", renk="#7C92AD", boyut=8.5, yatay="right", ic=0)
    ws.activate()


def ana_panel(wb, f):
    s = Sayfa(wb, "Ana Panel", f)
    kabuk(s, "Ana panel", "Ana panel", "Malzeme Stok  ›  Genel", "Tüm stokları güncelle")
    y = kpi_satiri(s, TOP + 16, [
        ("box", "#E0ECFB", "#1E6FD9", "Toplam stok değeri", "172,5 mn TL", "▼ %0,4", (SLBG, INK2), "düne göre", 0),
        ("cart", REDBG, RED, "Kritik stok", "37 kalem", "▼ 4", (GRNBG, "#166534"), "15,5 mn TL alım", 1),
        ("clock", AMBBG, AMB, "Atıl stok", "3,1 mn TL", "▲ %1,2", (REDBG, "#B91C1C"), "stokun %1,8'i", 2),
        ("swap", GRNBG, GRN, "Transfer fırsatı", "1,16 mn TL", "69 öneri", (SLBG, INK2), "bekliyor", 3)])
    y += 14
    lw = 740
    s.kart(X0, y, lw, 300)
    s.kutu(X0 + 16, y + 10, 300, 28, "Bugün yapılacaklar", renk=INK, boyut=10.5, kalin=True, ic=0)
    s.kutu(X0 + lw - 216, y + 10, 200, 28, "en önemliden başlayarak", renk=INK3, boyut=8.5, yatay="right", ic=0)
    isler = [("Satın almayı başlat: 37 malzemenin stoğu 1 aydan az yetiyor", "Satın alma", "15,5 mn TL", "Satın Alma"),
             ("69 stok naklini uygula (satın alma yerine)", "Transfer listesi", "1,16 mn TL", None),
             ("9 açık talep / siparişi iptal et: malzeme atıl stokta var", "Satın alma", "50,5 bin TL", "Satın Alma"),
             ("6 fazla teslim alınmış sipariş için iade kontrolü yap", "Satın alma", "—", "Satın Alma")]
    for i, (is_, nerede, tutar, hedef) in enumerate(isler):
        yy = y + 48 + i * 60
        url = f"internal:'{hedef}'!A1" if hedef else None
        if i:
            s.kutu(X0 + 16, yy - 4, lw - 32, 1, dolgu="#EEF2F6", ic=0)
        s.kutu(X0 + 16, yy + 10, 26, 26, str(i + 1), dolgu=REDBG if i == 0 else SLBG, renk="#B91C1C" if i == 0 else INK2,
               boyut=9, kalin=True, yatay="center", r=7, ic=0)
        s.kutu(X0 + 54, yy + 4, lw - 200, 22, is_, renk=INK, boyut=9.5, kalin=True, ic=0, url=url)
        s.kutu(X0 + 54, yy + 26, lw - 200, 18, nerede, renk=INK3, boyut=8.5, ic=0)
        s.kutu(X0 + lw - 136, yy + 10, 120, 26, tutar, renk=INK, boyut=10, kalin=True, yatay="right", ic=0)
    rx, rw = X0 + lw + 14, CW - lw - 14
    s.kart(rx, y, rw, 300)
    s.kutu(rx + 16, y + 10, 200, 28, "Stok dağılımı", renk=INK, boyut=10.5, kalin=True, ic=0)
    s.kutu(rx + rw - 176, y + 10, 160, 28, "13.11.2026 08:44", renk=INK3, boyut=8.5, yatay="right", ic=0)
    s.grafik(pasta(wb, 10, 4, ["#1E6FD9", "#19C2A8", "#F59E0B", VIO]), rx + 10, y + 40, rw - 20, 150)
    s.kutu(rx + 16, y + 192, 200, 24, "SAP bağlantıları", renk=INK, boyut=10, kalin=True, ic=0)
    s.kutu(rx + rw - 176, y + 192, 160, 24, "39 sn'de güncellendi", renk=INK3, boyut=8.5, yatay="right", ic=0)
    for i, ad in enumerate(["CDP / 7155", "EPP / 7055", "İlçeler", "Yüklenici"]):
        yy = y + 218 + i * 19
        s.kutu(rx + 16, yy + 5, 8, 8, dolgu=GRN, r=4, ic=0)
        s.kutu(rx + 32, yy, 200, 18, ad, renk=INK, boyut=9, ic=0)
        s.kutu(rx + rw - 96, yy, 80, 18, "Güncel", renk=GRN, boyut=9, kalin=True, yatay="right", ic=0)


STOK = [("7155", "D007", "Demontaj Ambarı", "ELEKTRİK MALZEMELERİ", "ÇIKMA/BDSZ", "1000001968", "36KV TIP1A (BOS)", 2, "ADT", 0.02),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000005638", "AKIM TRAFO 0010-0020/5-5A 1-3S EPOKSI D.", 1, "ADT", 780.00),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000005474", "AKIM TRAFO 0020/5A 5P10 15VA H TIP DENG", 1, "ADT", 901.80),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006195", "AKIM TRAFO 0150-300/5A TRODIAL", 3, "ADT", 13239.39),
        ("7155", "D007", "Demontaj Ambarı", "ELEKTRİK MALZEMELERİ", "ÇIKMA/BDSZ", "1000007438", "AKIM TRAFO 0300-600/5-5A 36kV TORODIAL", 9, "ADT", 0.09),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000007368", "AKIM TRF KABLO 300-600,5-5A 36KV,C0,5", 3, "ADT", 13669.86),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000007371", "AKIM TRF,36KV,KBL,S:0.5,100-200/5-5A", 6, "ADT", 34397.94),
        ("7155", "D003", "Müş.Bağ.Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006537", "AKIM,TRF,0.72KV,10VA,S:0.5,KABLO,100/5A", 87, "ADT", 15211.21),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006529", "AKIM,TRF,0.72KV,10VA,S:0.5,KABLO,1000/5A", 23, "ADT", 5639.60),
        ("7155", "D003", "Müş.Bağ.Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006529", "AKIM,TRF,0.72KV,10VA,S:0.5,KABLO,1000/5A", 15, "ADT", 3678.00),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006536", "AKIM,TRF,0.72KV,10VA,S:0.5,KABLO,150/5A", 153, "ADT", 23929.46),
        ("7155", "D003", "Müş.Bağ.Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006536", "AKIM,TRF,0.72KV,10VA,S:0.5,KABLO,150/5A", 72, "ADT", 11260.92),
        ("7155", "D001", "İşletme Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006535", "AKIM,TRF,0.72KV,10VA,S:0.5,KABLO,200/5A", 178, "ADT", 27089.59),
        ("7155", "D003", "Müş.Bağ.Ambarı", "ELEKTRİK MALZEMELERİ", "YENİ", "1000006535", "AKIM,TRF,0.72KV,10VA,S:0.5,KABLO,200/5A", 75, "ADT", 11414.15)]


def stoklar(wb, f):
    s = Sayfa(wb, "Stoklar", f)
    kabuk(s, "Tüm stoklar", "Tüm stoklar", "Malzeme Stok  ›  Stoklar", "Stokları güncelle", ["Excel", "Çıktı al"], sutunlar=[96, 70, 140, 176, 120, 108, 310, 70, 60, 90])
    y = kpi_satiri(s, TOP + 16, [
        ("clock", "#E0ECFB", "#1E6FD9", "Son güncelleme", "13.11.2026 08:44", None, None, "39 sn sürdü", None),
        ("list", SLBG, INK2, "Stok satırı", "734", None, None, "CDP 480  •  EPP 254", None),
        ("box", GRNBG, GRN, "Stok değeri", "72,7 mn TL", None, None, "72.683.606,98 TL", 0)])
    y += 14
    sekmeler(s, X0, y, ["Tümü · 734", "CDP · 480", "EPP · 254", "İlçe · 1.622", "Yüklenici · 474"])
    ust = y_satir(y + 34 + 10)
    tablo(s, ust, ["Üretim yeri", "Depo", "Depo adı", "Mal grubu", "Değerleme", "Malzeme", "Malzeme tanımı", "Miktar", "Birim", "Değer (TL)"],
          [(a, b, c, d, (e, "g" if e == "YENİ" else "s"), g, h, i, j, k) for a, b, c, d, e, g, h, i, j, k in STOK],
          [96, 70, 140, 176, 120, 108, 310, 70, 60, 90],
          ["tdm", "tdb", "td", "tdm", "st", "td", "td", "tdi", "tdm", "tdn"], suz=True)


def raporlar(wb, f, ngun):
    s = Sayfa(wb, "Raporlar", f)
    kabuk(s, "Raporlar", "Raporlar", "Malzeme Stok  ›  Rapor", None, ["Excel", "Çıktı al"], sutunlar=[120, 520, 180, 240, 180])
    y = kpi_satiri(s, TOP + 16, [
        ("box", "#E0ECFB", "#1E6FD9", "CDP stok değeri", "20,0 mn TL", None, None, "480 satır  •  9 depo", None),
        ("box", "#E6FAF6", TEAL_D, "EPP stok değeri", "52,7 mn TL", None, None, "254 satır  •  5 depo", None),
        ("pin", "#EDE9FE", VIO, "İlçe stok değeri", "6,3 mn TL", None, None, "1.622 satır  •  16 depo", None),
        ("people", AMBBG, AMB, "Yüklenici stok değeri", "93,5 mn TL", None, None, "474 satır  •  4 depo", None)])
    y += 14
    w3 = (CW - 28) // 3
    basliklar = ["Depolara göre CDP stok dağılımı (mn TL)", "İllere göre stok değeri (mn TL)", "Stok değeri eğilimi (mn TL)"]
    grafikler = [pasta(wb, 16, 6, ["#1E6FD9", "#19C2A8", "#F59E0B", VIO, "#EC4899", "#94A3B8"]),
                 sutun(wb, 24, 5, 1, "#1E6FD9", "0"), cizgi(wb, 32, ngun, 1, TEAL_D)]
    for i in range(3):
        x = X0 + i * (w3 + 14)
        s.kart(x, y, w3, 280)
        s.kutu(x + 16, y + 10, w3 - 32, 28, basliklar[i], renk=INK, boyut=10, kalin=True, ic=0)
        s.grafik(grafikler[i], x + 8, y + 42, w3 - 16, 230)
    y += 280 + 14
    ust = y_satir(y)
    depolar = [("D001", "İşletme Ambarı", 320, 8075980.79, 0.404), ("D002", "Yatırım Ambarı", 5, 13335.11, 0.001),
               ("D003", "Müş.Bağ.Ambarı", 28, 11489464.05, 0.574), ("D005", "Sayaç Dm Ambar", 24, 292124.86, 0.015),
               ("D010", "İSG-HIRDAVAT-MÜŞ", 14, 91812.01, 0.005), ("D011", "PLANLI BAKIM S3", 7, 44862.51, 0.002)]
    tablo(s, ust, ["Depo", "Depo adı", "Stok satırı", "Stok değeri (TL)", "Pay"], depolar, [120, 520, 180, 240, 180],
          ["tdb", "td", "tdi", "tdn", "tdp"])
    s.ws.conditional_format(ust + 1, 6, ust + len(depolar), 6, {"type": "data_bar", "bar_color": "#9CC3F5", "bar_solid": True})


ATIL = [("Trabzon", "KABLO YE3SV 1X240/25MM2 36KV CU", "D002", 478.01, "M", 719485.06, ("Kendi ihtiyacında kullan", "r"), "30.06.2024"),
        ("Giresun", "DIREK AYD.GALV.AD1(POL) 100/10 NORM.TIP", "D002", 4, "ADT", 17255.42, ("Kendi ihtiyacında kullan", "r"), "31.12.2023"),
        ("Gümüşhane", "ARMATUR LED TEDAS MYD/SINIF 10", "D002", 4, "ADT", 9275.49, ("Kendi ihtiyacında kullan", "r"), "31.05.2024"),
        ("Trabzon", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", "D001", 2, "ADT", 151527.70, ("Transfer et → Giresun", "b"), "07.02.2024"),
        ("Trabzon", "KABLO NYM 4X16MM2", "D002", 600, "M", 96016.50, ("Transfer et → Gümüşhane", "b"), "05.11.2024"),
        ("Rize", "KABLO ALPEK 3X95+16+95MM2", "D001", 272, "M", 68231.25, ("Transfer et → Trabzon", "b"), "02.10.2024"),
        ("Giresun", "KABLO YE3SV 1X240/25MM2 36KV CU", "D002", 88, "M", 39446.85, ("Transfer et → Trabzon", "b"), "31.12.2023"),
        ("Rize", "YUK AYR.SIG 160A 00BOY DIKEY", "D001", 9, "ADT", 7116.14, ("Önce bunu kullan", "a"), "07.11.2024"),
        ("Artvin", "SAYAC TRIFAZE BAYLAN", "D005", 241, "ADT", 8932.13, ("Önce bunu kullan", "a"), "31.12.2023"),
        ("Giresun", "YAGMURLUK-KAPSONLU,PVC,EDAS CORUH,KIS", "D009", 1, "ADT", 752.76, ("Beklet", "s"), "23.10.2024"),
        ("Artvin", "SAYAC TRIFAZE GSM KESME (LUNA)", "D005", 1, "ADT", 720.00, ("Sat / hurda / iade", "s"), "31.12.2023"),
        ("Trabzon", "KURMA MOTORU 220V (ULUSOY)", "D001", 2, "ADT", 800.00, ("Sat / hurda / iade", "s"), "31.12.2023")]


def atil(wb, f):
    s = Sayfa(wb, "Atıl Stok", f)
    kabuk(s, "Atıl stok", "Atıl stok", "Malzeme Stok  ›  Karar", "SAP'den oku", ["Excel", "Transfer listesi"], sutunlar=[130, 400, 80, 100, 70, 130, 220, 110])
    y = kpi_satiri(s, TOP + 16, [
        ("clock", REDBG, RED, "Atıl (24 ay+)", "31,9 mn TL", "1.506 kalem", (SLBG, INK2), None, 4),
        ("clock", AMBBG, AMB, "3 ay içinde atıl", "2,9 mn TL", "336 kalem", (SLBG, INK2), None, 5),
        ("swap", GRNBG, GRN, "Kullanılabilir", "1,44 mn TL", "110 kalem", (SLBG, INK2), "kendi + transfer", 6),
        ("box", SLBG, INK2, "Sat / hurda adayı", "21,1 mn TL", "735 kalem", (SLBG, INK2), None, 7)])
    y += 14
    x = sekmeler(s, X0, y, ["Ne yapılmalı · 2.130", "Özet", "Atıl listesi · 1.798"])
    s.kutu(X0 + CW - 420, y + 4, 420, 26, "SAP okuması 04.11.2026  •  7 gün önce: yeniden okuyun", renk=RED, boyut=9, kalin=True, yatay="right", ic=0)
    ust = y_satir(y + 34 + 10)
    tablo(s, ust, ["İl", "Malzeme", "Depo", "Miktar", "Birim", "Değer (TL)", "Ne yapılmalı", "En eski giriş"], ATIL,
          [130, 400, 80, 100, 70, 130, 220, 110], ["tdb", "td", "tdm", "tdq", "tdm", "tdn", "st", "tdm"], suz=True)


def satinalma(wb, f):
    s = Sayfa(wb, "Satın Alma", f)
    kabuk(s, "Satın alma", "Satın alma", "Malzeme Stok  ›  Karar", "SAP'den oku", ["Excel", "Transfer listesi"], sutunlar=[130, 400, 130, 100, 100, 70, 150, 160])
    y = kpi_satiri(s, TOP + 16, [
        ("cart", REDBG, RED, "Satın alınması gereken", "15,5 mn TL", "37 kalem", (REDBG, "#B91C1C"), "kritik stok", 1),
        ("doc", AMBBG, AMB, "Gereksiz talep / sipariş", "50,5 bin TL", "9 kalem", (SLBG, INK2), "atıl stokta var", None),
        ("swap", GRNBG, GRN, "Transferle çözülen", "1,16 mn TL", "69 kalem", (SLBG, INK2), None, 3),
        ("box", "#EDE9FE", VIO, "Fazla teslim alınan", "6 kalem", None, None, "iade / fatura kontrolü", None)])
    y += 14
    sekmeler(s, X0, y, ["Satın alınacaklar · 37", "İptal / azalt · 9", "Fazla teslim · 6"])
    ust = y_satir(y + 34 + 10)
    veri = [("Artvin EPP", "BARET ELEKTRIKCI 30 KV ROCKMAN", 0, 0, 7, "ADT", 3046.64, ("Stok yok", "r")),
            ("Artvin CDP", "DEDEKTOR AG SES ISIK TITRESIM 20-1000V", 0, 0, 11, "ADT", 14913.08, ("Stok yok", "r")),
            ("Artvin EPP", "PANTOLON-EDAS CORUH YAZ", 0, 0, 136, "ADT", 82960.00, ("Stok yok", "r")),
            ("Artvin EPP", "SWEATSHIRT-UZUN KOL POLO CORUH EDAS YAZ", 3, 8, 213, "ADT", 64331.32, ("3 gün yeter", "r")),
            ("Artvin CDP", "KABLO ALPEK 3X50+16+70MM2", 3, 10, 261.5, "M", 40766.80, ("3 gün yeter", "r")),
            ("Artvin CDP", "KLEMENS AL-CU A40-B25", 8, 12, 128, "ADT", 5636.37, ("8 gün yeter", "r")),
            ("Artvin CDP", "KABLO ALPEK 3X16+16+25MM2", 23, 110.65, 318.35, "M", 18384.40, ("23 gün yeter", "r")),
            ("Artvin CDP", "KABLO ALPEK 3X25+16+35MM2", 24, 522.9, 1424.53, "M", 127506.20, ("24 gün yeter", "r")),
            ("Artvin CDP", "KLEMENS AL A40-A40", 32, 92, 167, "ADT", 3619.11, ("1 ay 2 gün yeter", "a")),
            ("Artvin EPP", "DIREK DEMIR GAL KAY G-K2", 38, 39, 37, "ADT", 1054924.61, ("1 ay 8 gün yeter", "a")),
            ("Artvin CDP", "KABLO ALPEK 2X16+25MM2", 41, 790.61, 935.87, "M", 41096.12, ("1 ay 11 gün yeter", "a")),
            ("Artvin CDP", "TRAVERS DEMIR GAL OG GN-250", 63, 48, 21, "ADT", 58954.40, ("2 ay 3 gün yeter", "a"))]
    tablo(s, ust, ["İl / sistem", "Malzeme", "Yeter (gün)", "Stok", "Alınacak", "Birim", "Tahmini tutar (TL)", "Durum"], veri,
          [130, 400, 130, 100, 100, 70, 150, 160], ["tdb", "td", "tdi", "tdq", "tdq", "tdm", "tdn", "st"], suz=True)
    s.ws.conditional_format(ust + 1, 4, ust + len(veri), 4, {"type": "data_bar", "bar_color": "#FCA5A5", "bar_solid": True, "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 90})


def arama(wb, f):
    s = Sayfa(wb, "Malzeme Arama", f)
    kabuk(s, "Malzeme arama", "Malzeme arama", "Malzeme Stok  ›  Genel", "Bütün illeri güncelle", sutunlar=[70, 70, 160, 110, 310, 130, 70, 60, 110, 80, 70])
    # Arama bloğu hücrelerden oluşur (şekil hücrenin üstünü kapatacağı için kart şekli kullanılmaz):
    # beyaz satırlar + yazılabilir arama hücresi + "Ara" düğmesi.
    ws = s.ws
    for r in (4, 5, 6):
        ws.set_row_pixels(r, 30)
        for c in range(2, 13):
            ws.write_blank(r, c, None, f["ak"])
    ws.write(4, 2, "Malzeme kodu ya da açıklaması", f["etiket"])
    ws.merge_range(5, 2, 5, 9, "Trafo 50kva", f["giris"])
    s.kutu(X0 + CW - 150, 94 + 26, 134, 38, "Ara", dolgu=NAVY, renk=WHITE, boyut=10.5, kalin=True, yatay="center", r=8)
    y = 94 + 90 + 12
    x = X0
    for i, il in enumerate(["Tümü", "Giresun", "Trabzon", "Rize", "Gümüşhane", "Artvin"]):
        w = 30 + 8 * len(il)
        s.kutu(x, y, w, 30, il, dolgu=NAVY if i == 5 else WHITE, cizgi=None if i == 5 else LINE, renk=WHITE if i == 5 else INK2,
               boyut=9, kalin=True, yatay="center", r=15)
        x += w + 8
    s.kutu(X0 + CW - 520, y, 520, 30, "15 satır  •  10 depo  •  22 adet  •  477.109,92 TL", renk=INK, boyut=10, kalin=True, yatay="right", ic=0)
    ust = y_satir(y + 30 + 10)
    veri = [("7155", "D006", "Arıza Mlz. Ambar", "1000007403", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", ("ÇIKMA/BDSZ", "s"), 1, "ADT", 0.01, "CDP", "Artvin"),
            ("7155", "D007", "Demontaj Ambarı", "1000007403", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", ("ÇIKMA/BDSZ", "s"), 4, "ADT", 0.04, "CDP", "Artvin"),
            ("7055", "D002", "Yatırım Ambarı", "1000007403", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", ("YENİ", "g"), 4, "ADT", 318073.18, "EPP", "Artvin"),
            ("7155", "K001", "Ardanuç KA", "1000005009", "TRAFO OG/AG 0050kVA 6,3/0,4KV(GEN.DEP)", ("ÇIKMA/BDSZ", "s"), 1, "ADT", 0.01, "İlçe", "Artvin"),
            ("7155", "K008", "Şavşat KA", "1000007403", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", ("ÇIKMA/BDSZ", "s"), 1, "ADT", 0.01, "İlçe", "Artvin"),
            ("7155", "K010", "Yusufeli KA", "1000007403", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", ("ÇIKMA/BDSZ", "s"), 1, "ADT", 0.01, "İlçe", "Artvin"),
            ("7055", "M123", "26.GR.02 ENGİNDE", "1000007403", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", ("YENİ", "g"), 1, "ADT", 79518.29, "Yüklenici", "Artvin"),
            ("7055", "M124", "26.KET.01 ENER", "1000007403", "TRAFO OG/AG 0050kVA 33/0,4kV(HERM)", ("YENİ", "g"), 1, "ADT", 79518.29, "Yüklenici", "Artvin")]
    tablo(s, ust, ["ÜY", "Depo", "Depo tanımı", "Malzeme", "Malzeme açıklaması", "Parti", "Miktar", "Birim", "Değer (TL)", "Kaynak", "İl"], veri,
          [70, 70, 160, 110, 310, 130, 70, 60, 110, 80, 70], ["tdm", "tdb", "td", "td", "td", "st", "tdi", "tdm", "tdn", "tdm", "tdb"], suz=True)



def yonetim(wb, f, ngun):
    s = Sayfa(wb, "Yönetim Merkezi", f)
    kabuk(s, "Yönetim merkezi", "Yönetim merkezi", "Çoruh EDAŞ  ›  Tüm iller", "Tüm birimleri güncelle", ["PDF özeti", "Web panosu"], sutunlar=[200, 240, 200, 160, 220, 220])
    y = kpi_satiri(s, TOP + 16, [
        ("shield", GRNBG, GRN, "Sağlık skoru", "64 / 100", "▲ 3 puan", (GRNBG, "#166534"), "düne göre", None),
        ("box", "#E0ECFB", "#1E6FD9", "Tekil stok değeri", "1,39 mr TL", "▼ 2,3 mn", (SLBG, INK2), None, 0),
        ("clock", AMBBG, AMB, "Atıl stok", "31,9 mn TL", "▼ %0,6", (GRNBG, "#166534"), "stokun %2,3'ü", 4),
        ("cart", REDBG, RED, "Kritik stok", "179 kalem", "44 stoksuz", (REDBG, "#B91C1C"), "79,5 mn TL alım", 1)])
    y += 14
    lw = 600
    s.kart(X0, y, lw, 320)
    s.kutu(X0 + 16, y + 10, 300, 28, "Bugün yapılacak 5 iş", renk=INK, boyut=10.5, kalin=True, ic=0)
    isler = [("Satın almayı başlat: 179 malzemenin stoğu 1 aydan az yetiyor", "79,5 mn TL"),
             ("Artvin: 834 adet İSG malzemesi zimmetlenmemiş", "—"),
             ("286 stok naklini / atıl kullanımını uygula", "7,6 mn TL"),
             ("76 açık talep / siparişi iptal et ya da azalt", "1,8 mn TL"),
             ("41 fazla teslim için iade / fatura kontrolü yap", "2,1 mn TL")]
    for i, (is_, tutar) in enumerate(isler):
        yy = y + 46 + i * 52
        if i:
            s.kutu(X0 + 16, yy - 3, lw - 32, 1, dolgu="#EEF2F6", ic=0)
        s.kutu(X0 + 16, yy + 8, 26, 26, str(i + 1), dolgu=REDBG if i == 0 else SLBG, renk="#B91C1C" if i == 0 else INK2,
               boyut=9, kalin=True, yatay="center", r=7, ic=0)
        s.kutu(X0 + 54, yy + 8, lw - 190, 26, is_, renk=INK, boyut=9.5, kalin=True, ic=0)
        s.kutu(X0 + lw - 126, yy + 8, 110, 26, tutar, renk=INK, boyut=10, kalin=True, yatay="right", ic=0)
    w2 = (CW - lw - 28) // 2
    for i, (baslik, ch) in enumerate([("İllere göre stok değeri (mn TL)", sutun(wb, 24, 5, 1, "#1E6FD9", "0")),
                                      ("Sağlık skoru eğilimi", cizgi(wb, 32, ngun, 2, TEAL_D))]):
        x = X0 + lw + 14 + i * (w2 + 14)
        s.kart(x, y, w2, 320)
        s.kutu(x + 16, y + 10, w2 - 32, 28, baslik, renk=INK, boyut=10, kalin=True, ic=0)
        s.grafik(ch, x + 8, y + 42, w2 - 16, 270)
    y += 320 + 14
    ust = y_satir(y)
    veri = [("Artvin", 172.5e6, 3.1e6, 0.018, 2, ("Demontaj girişi yapılmalı", "r")), ("Trabzon", 498.2e6, 9.8e6, 0.020, 9, ("Demontaj girişi yapılmalı", "r")),
            ("Gümüşhane", 141.9e6, 4.6e6, 0.032, 0, ("Tamam", "g")), ("Giresun", 318.4e6, 8.9e6, 0.028, 5, ("Demontaj girişi yapılmalı", "r")),
            ("Rize", 258.1e6, 5.5e6, 0.021, 3, ("Demontaj girişi yapılmalı", "r"))]
    tablo(s, ust, ["İl", "Tekil stok değeri (TL)", "Atıl stok (TL)", "Atıl oranı", "Demontaj girişi eksik kalem", "Durum"], veri,
          [200, 240, 200, 160, 220, 220], ["tdb", "tdn", "tdn", "tdp", "tdi", "st"])


def kapsam(wb, f):
    ws = wb.add_worksheet("Kapsam")
    ws.hide_gridlines(2)
    ws.set_column_pixels(0, 0, 24)
    ws.set_column_pixels(1, 1, 220)
    ws.set_column_pixels(2, 2, 760)
    ws.set_column_pixels(3, 3, 420)
    b = wb.add_format({"font_name": FONT, "font_size": 14, "bold": True, "font_color": INK})
    n = wb.add_format({"font_name": FONT, "font_size": 10, "font_color": INK2, "text_wrap": True, "valign": "top"})
    ws.write(1, 1, "AKSORA v2  —  hangi şablon hangi sayfaları kapsar", b)
    ws.merge_range(2, 1, 2, 3, "Bu dosya makrosuz bir görünüm prototipidir; rakamlar örnektir. Soldaki menü ve üst banttaki düğmeler "
                   "sayfalar arasında gezinir. Dosyadaki her öğe hücre, şekil, grafik ya da köprüdür.", n)
    ws.set_row_pixels(2, 40)
    ws.write_row(4, 1, ["Şablon (bu dosyada)", "AKSORA'daki sayfalar", "Not"], f["th"])
    satirlar = [
        ("Giriş", "BIRIM SECIM, KAPAK, YATIRIM / MUSTERI / ILCE IL SECIM", "İl seçimi üst banttaki il düğmesine taşınır."),
        ("Ana Panel", "ANA PANEL, M_ANA PANEL, Y_ANA PANEL, I_ANA PANEL", "Alt sıradaki sayfa düğmeleri sol menüye taşınır."),
        ("Stoklar", "STOKLAR, CDP_STOK, EPP_STOK, İLÇE_STOK, YÜKLENİCİ_STOK, M_ / Y_ / I_ stok sayfaları (17 sayfa)", "Tek şablon; sekmeler stok türünü seçer."),
        ("Raporlar", "RAPORLAR, CDP_RAPOR, EPP_RAPOR, İLÇE_RAPORLAR, YÜKLENİCİ_RAPOR, M_ / Y_ / I_ rapor sayfaları (15 sayfa)", "Grafikler Excel'in kendi grafikleridir."),
        ("Atıl Stok", "ATIL_STOK, TRANSFER_LISTESI, STOK_KARŞILAŞTIRMA, I_STOK_KARŞILAŞTIRMA", "Karar listesi kalıbı: göstergeler, sekmeler, durum sütunu."),
        ("Satın Alma", "SATINALMA_KONTROL, Y_TEMİNAT, Y_HAKEDİŞ, I_AYLIK_DEMONTAJ, I_İSG_MALZEMELERİ", "Aynı kalıp; veri çubukları koşullu biçimlendirmedir."),
        ("Malzeme Arama", "MALZEME_ARA, M_MALZEME_ARA, Y_MALZEME_ARA, I_MALZEME_ARA", "Arama kutusu gerçek, yazılabilir bir hücredir."),
        ("Yönetim Merkezi", "ADMIN PANEL", "Karne, işler ve il karşılaştırması aynı kabukta."),
    ]
    td = wb.add_format({"font_name": FONT, "font_size": 10, "font_color": INK, "text_wrap": True, "valign": "top", "bottom": 1, "bottom_color": "#F0F3F7"})
    for i, (a, c, d) in enumerate(satirlar):
        ws.set_row_pixels(5 + i, 40)
        ws.write_url(5 + i, 1, f"internal:'{a}'!A1", td, a)
        ws.write(5 + i, 2, c, td)
        ws.write(5 + i, 3, d, td)


# ---------------------------------------------------------------- şekil XML'i: yuvarlak köşe, gölge, iç boşluk
GOLGE = ('<a:effectLst><a:outerShdw blurRad="76200" dist="19050" dir="5400000" algn="t" rotWithShape="0">'
         '<a:srgbClr val="0F172A"><a:alpha val="14000"/></a:srgbClr></a:outerShdw></a:effectLst>')


def sekilleri_isle(yol):
    gecici = yol + ".tmp"
    with zipfile.ZipFile(yol) as zi, zipfile.ZipFile(gecici, "w", zipfile.ZIP_DEFLATED) as zo:
        for ad in zi.namelist():
            veri = zi.read(ad)
            if ad.startswith("xl/drawings/drawing") and ad.endswith(".xml"):
                veri = _drawing(veri.decode("utf-8")).encode("utf-8")
            zo.writestr(ad, veri)
    shutil.move(gecici, yol)


def _drawing(xml):
    def sp(m):
        blok = m.group(0)
        e = re.search(r'descr="r=(\d+);g=(\d);i=(\d+)"', blok)
        if not e:
            return blok
        r, g, i = int(e.group(1)), e.group(2) == "1", int(e.group(3))
        en = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', blok)
        if r and en:
            kisa = min(int(en.group(1)), int(en.group(2))) / 9525.0
            oran = max(0, min(50000, int(100000 * r / max(kisa, 1))))
            blok = blok.replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
                                f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {oran}"/></a:avLst></a:prstGeom>')
        if g:
            blok = blok.replace("</xdr:spPr>", GOLGE + "</xdr:spPr>")
        ins = i * 9525
        blok = re.sub(r"<a:bodyPr ", f'<a:bodyPr lIns="{ins}" rIns="{ins}" tIns="0" bIns="0" ', blok, count=1)
        return blok.replace(e.group(0), 'descr=""')
    xml = re.sub(r"<xdr:sp .*?</xdr:sp>", sp, xml, flags=re.S)
    # Katman sırası: XlsxWriter grafikleri en alta yazar; kartların üstünde görünmeleri için sona taşınır.
    capa = re.compile(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", re.S)
    parcalar = [m.group(0) for m in capa.finditer(xml)]
    if parcalar:
        bas = xml[:capa.search(xml).start()]
        son = xml[xml.rfind("</xdr:" + re.search(r"<xdr:(\w+Anchor)", parcalar[-1]).group(1) + ">") + len("</xdr:" + re.search(r"<xdr:(\w+Anchor)", parcalar[-1]).group(1) + ">"):]
        sekil = [x for x in parcalar if "<xdr:graphicFrame" not in x]
        grafik = [x for x in parcalar if "<xdr:graphicFrame" in x]
        xml = bas + "".join(sekil) + "".join(grafik) + son
    return xml


def main(yol):
    wb = xlsxwriter.Workbook(yol)
    f = bicimler(wb)
    giris(wb, f)
    ana_panel(wb, f)
    stoklar(wb, f)
    ngun = 15
    raporlar(wb, f, ngun)
    atil(wb, f)
    satinalma(wb, f)
    arama(wb, f)
    yonetim(wb, f, ngun)
    kapsam(wb, f)
    veri_sayfasi(wb)
    wb.close()
    sekilleri_isle(yol)
    print("yazıldı:", yol)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "AKSORA_v2_Prototip.xlsx")
