#!/usr/bin/env python3
"""AKSORA VBA araçları (Excel gerektirmez).

  cikar   <AKSORA.xlsm> <hedef_klasor> [--maskele]
          Çalışma kitabındaki VBA kodunu UTF-8 metin dosyalarına çıkarır
          (sürüm takibi ve inceleme için). --maskele: şifre sabitlerini gizler.

  bagla   <kaynak_klasor> <baglantilar.json> [--yaz]
          Geliştirme paketinin bağlantı satırlarını kaynak üzerinde dener.
          Her çapa tam olarak bir kez bulunmalıdır. --yaz: dosyaları günceller.

  sablon  <pano_sablon.html> <modAKSORA_PANO_SABLON.bas>
          Web panosu şablonunu VBA metin modülüne çevirir ve geri okuyarak doğrular.

  denetle <kaynak_klasor>
          Statik denetim özeti (satır sayıları, riskli kalıplar, düz metin şifreler).

Gereken: pip install oletools  (yalnız "cikar" için)
"""
import collections
import glob
import json
import os
import re
import sys

SIFRE_SABITI = re.compile(r'(Const\s+\w*SIFRE\w*\s+As\s+String\s*=\s*)"[^"]*"', re.I)


def cikar(xlsm, hedef, maskele):
    from oletools.olevba import VBA_Parser

    os.makedirs(hedef, exist_ok=True)
    ayristirici = VBA_Parser(xlsm)
    adet = satir = 0
    for _, _, ad, kod in ayristirici.extract_macros():
        if ad.lower().endswith(".txt"):  # p-code dökümü
            continue
        kod = kod.replace("\r\n", "\n")
        if maskele:
            kod = SIFRE_SABITI.sub(r'\1"***"', kod)
        with open(os.path.join(hedef, ad.replace("/", "_")), "w", encoding="utf-8", newline="\n") as f:
            f.write(kod)
        adet += 1
        satir += kod.count("\n")
    print(f"{adet} bileşen, {satir} satır -> {hedef}")


def _modul_yolu(klasor, modul):
    for uzanti in (".bas", ".cls", ".frm"):
        yol = os.path.join(klasor, modul + uzanti)
        if os.path.exists(yol):
            return yol
    return None


def baglanti_uygula(satirlar, b):
    """Kurulum betiğiyle (PowerShell) aynı kural. Dönüş: (durum, yeni_satirlar)."""
    hedef = b["satir"].strip().lower()
    if any(s.strip().lower() == hedef for s in satirlar):
        return "zaten var", satirlar
    capa = b["capa"].strip().lower()
    if b["eslesme"] == "tam":
        bulunan = [i for i, s in enumerate(satirlar) if s.strip().lower() == capa]
    else:
        bulunan = [i for i, s in enumerate(satirlar) if s.strip().lower().startswith(capa)]
    if len(bulunan) != 1:
        return f"HATA: çapa {len(bulunan)} kez bulundu", satirlar
    i = bulunan[0]
    if satirlar[i].rstrip().endswith(" _"):
        return "HATA: çapa satırı devam satırı", satirlar
    if b["konum"] == "once" and i > 0 and satirlar[i - 1].rstrip().endswith(" _"):
        return "HATA: çapa bir devam satırının parçası", satirlar
    yer = i + 1 if b["konum"] == "sonra" else i
    return f"eklendi (satır {yer + 1})", satirlar[:yer] + [b["satir"]] + satirlar[yer:]


def bagla(klasor, json_yolu, yaz):
    with open(json_yolu, encoding="utf-8") as f:
        baglantilar = json.load(f)
    hata = 0
    for b in baglantilar:
        yol = _modul_yolu(klasor, b["modul"])
        if not yol:
            print(f"HATA: {b['modul']} bulunamadı")
            hata += 1
            continue
        with open(yol, encoding="utf-8") as f:
            satirlar = f.read().split("\n")
        durum, yeni = baglanti_uygula(satirlar, b)
        print(f"{b['modul']:22s} {durum:28s} {b['satir'].strip()[:70]}")
        if durum.startswith("HATA"):
            hata += 1
        elif yaz and yeni is not satirlar:
            with open(yol, "w", encoding="utf-8", newline="\n") as f:
                f.write("\n".join(yeni))
    return hata


def denetle(klasor):
    dosyalar = sorted(glob.glob(os.path.join(klasor, "*.bas")) + glob.glob(os.path.join(klasor, "*.cls")) +
                      glob.glob(os.path.join(klasor, "*.frm")))
    kaliplar = {
        "On Error Resume Next": r"On Error Resume Next",
        ".Select / .Activate": r"\.(Select|Activate)\b",
        "MsgBox": r"\bMsgBox\b",
        "Declare (Windows API)": r"^\s*(Private |Public )?Declare ",
        "VBScript.RegExp": r"VBScript\.RegExp",
        "Düz metin şifre sabiti": r'Const\s+\w*SIFRE\w*\s+As\s+String\s*=\s*"[^"*]',
        'Parolasız koruma (Password:="")': r'Password:=""',
    }
    sayac = collections.Counter()
    toplam = prosedur = explicit_yok = 0
    buyuk = []
    for yol in dosyalar:
        with open(yol, encoding="utf-8") as f:
            kod = f.read()
        n = kod.count("\n")
        toplam += n
        prosedur += len(re.findall(r"^(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function|Property \w+) ", kod, re.M))
        if n > 20 and "Option Explicit" not in kod:
            explicit_yok += 1
        buyuk.append((n, os.path.basename(yol)))
        for ad, kalip in kaliplar.items():
            sayac[ad] += len(re.findall(kalip, kod, re.M | re.I))
    print(f"Bileşen: {len(dosyalar)}  |  Satır: {toplam}  |  Prosedür: {prosedur}  |  Option Explicit eksik: {explicit_yok}")
    print("En büyük modüller: " + ", ".join(f"{ad} ({n})" for n, ad in sorted(buyuk, reverse=True)[:5]))
    for ad in kaliplar:
        print(f"  {ad:34s} {sayac[ad]}")


def sablon(html_yolu, bas_yolu):
    with open(html_yolu, encoding="utf-8") as f:
        metin = f.read().replace("\r\n", "\n")
    metin.encode("cp1254")  # VBA düzenleyicisi Windows-1254 bekler; uymayan karakter varsa burada durur
    satirlar = metin.split("\n")
    if satirlar and satirlar[-1] == "":
        satirlar.pop()
    parca_boyu, satir_boyu = 40, 700
    govde, cagri = [], []
    for no, bas in enumerate(range(0, len(satirlar), parca_boyu), 1):
        cagri.append(f"PS_{no:02d}()")
        govde.append(f"Private Function PS_{no:02d}() As String")
        govde.append("    Dim s As String")
        for satir in satirlar[bas:bas + parca_boyu]:
            dilimler = [satir[i:i + satir_boyu] for i in range(0, len(satir), satir_boyu)] or [""]
            for i, d in enumerate(dilimler):
                son = " & vbLf" if i == len(dilimler) - 1 else ""
                govde.append('    s = s & "' + d.replace('"', '""') + '"' + son)
        govde.append(f"    PS_{no:02d} = s")
        govde.append("End Function")
        govde.append("")
    bas_metin = ['Attribute VB_Name = "modAKSORA_PANO_SABLON"', "Option Explicit", "",
                 "' OTOMATIK URETILIR - ELLE DUZENLEMEYIN.",
                 "' Kaynak: gelistirme/pano/pano_sablon.html  |  Uretim: python3 tools/aksora_vba.py sablon ...", "",
                 "Public Function AKSORA_PANO_SABLON() As String"]
    for i in range(0, len(cagri), 6):
        bas_metin.append("    AKSORA_PANO_SABLON = AKSORA_PANO_SABLON & " + " & ".join(cagri[i:i + 6]))
    bas_metin += ["End Function", ""]
    cikti = "\n".join(bas_metin + govde)
    with open(bas_yolu, "w", encoding="utf-8", newline="\n") as f:
        f.write(cikti)
    # Geri okuma: VBA metin sabitleri birleştirilince şablonun aynısı çıkmalı.
    geri = ""
    for m in re.finditer(r'^    s = s & "((?:[^"]|"")*)"( & vbLf)?$', cikti, re.M):
        geri += m.group(1).replace('""', '"') + ("\n" if m.group(2) else "")
    if geri != "\n".join(satirlar) + "\n":
        print("HATA: geri okunan metin şablonla aynı değil")
        return 1
    en_uzun = max(len(s) for s in cikti.split("\n"))
    print(f"{len(satirlar)} satır, {len(cagri)} parça, en uzun VBA satırı {en_uzun} -> {bas_yolu} (geri okuma doğrulandı)")
    return 0


def main(a):
    if len(a) >= 3 and a[0] == "cikar":
        cikar(a[1], a[2], "--maskele" in a)
    elif len(a) >= 3 and a[0] == "bagla":
        sys.exit(1 if bagla(a[1], a[2], "--yaz" in a) else 0)
    elif len(a) >= 3 and a[0] == "sablon":
        sys.exit(sablon(a[1], a[2]))
    elif len(a) >= 2 and a[0] == "denetle":
        denetle(a[1])
    else:
        print(__doc__)
        sys.exit(2)


if __name__ == "__main__":
    main(sys.argv[1:])
