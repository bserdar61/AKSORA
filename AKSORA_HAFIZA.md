# AKSORA – Sistem Hafızası (kilitli temel, 13.10.2026)

Bu dosya, AKSORA'nın kilitlenmiş tasarım ve çalışma kurallarıdır. Yeni il / yeni birim
eklenirken bu kurallar birebir uygulanır; "aynı olanlara dokunulmaz".

## Kullanıcı ve çalışma şekli
- Burhan Serdar, Çoruh EDAŞ malzeme stok uzmanı. Türkçe, kısa, açıklamasız; "ne yapacağım" net söylenir.
- VBA proje şifresi: kullanıcıda (yalnız VBE kilidi için).
- Temel dosya: GitHub `main` dalındaki kullanıcının son yüklediği `AKSORA.xlsm`.
  Yeni dosya her zaman bu dosyanın üzerine kurulur (veriler korunur). Kod değişiklikleri
  `claude/aksora-vba-error-fix-a7b5so` dalında `vba/*.bas` olarak tutulur.

## Derleme (Linux)
- Kaynaklar: scratchpad `raw/*.txt` (cp1254, CRLF). Değişiklik → `tools/vbalint.py` →
  `tools/build_bin.py <temel vbaProject.bin> raw new_vbaProject.bin` →
  `tools/build_base.py <temel xlsm> new_vbaProject.bin final/AKSORA.xlsm`.
- Yeni modül eklenmez; yeni kod mevcut modüllere eklenir.

## Kilitli şablonlar
1. **Ana panel = Malzeme biriminin ANA PANEL'i** (4 kart: başlık lacivert, DURUM / SON GÜNCELLEME /
   STOK SATIRI / STOK DEĞERİ / DEPO SAYISI; LED: güncellenirken mavi, güncel yeşil, 1 gün turuncu,
   2+ gün kırmızı; durum şeridi solda: "Son güncelleme | Güncelleme süresi X sn | grup satırları |
   Toplam stok değeri", güncellenirken açık sarı "Güncelleniyor • adım"; sağda hap
   "⇄ İL DEĞİŞTİR" (AKSORA_IL_DEGISTIR_STIL); alt lacivert bant + altında açık "Panel kontrolü" şeridi).
2. **Stok sayfası = Müşteri M_CDP_STOK** (tam genişlik, KPI kartları eşit yayılır, depo kartı + açılır
   liste, butonlar sağa hizalı, başlıklar: Üretim Yeri, Depo, Depo Adı, Mal Grubu, …, Seri Numarası;
   filtreye uyumlu beyaz/mavi bant, tablo son satırda biter, TL biçimi).
3. **Rapor = Müşteri M_CDP_RAPOR** (01/02/03 kartlar; tablo: Depo, Depo Adı, Stok Satırı, Seri Numaralı,
   Toplam Stok Değeri + TOPLAM; sağda "DEPOLARA GÖRE STOK DAĞILIMI" halka grafik + PAY listesi;
   kart 1 tablo genişliğinde, kart 2-3 ve grafik sağ kenara kadar; alt not; logo sağ altta).

## Kilitli çalışma sistemi
- SAP okuma: Excel açılmaz. ALV listesi "Yerel dosya…" (MB52 btn[45]) → "Sekmeli metin" →
  `C:\TEMP` (yoksa oluşturulur, olmazsa %TEMP%) → okunur → silinir. Pencere alanları sabit
  kimliklerle doldurulur (`wnd[1]/usr/ctxtDY_PATH`, `DY_FILENAME`, `DY_FILE_ENCODING`=4110).
  Olmazsa eski yönteme düşülür; toplu güncellemede kullanıcıya sorulur.
- Müşteri: IQ09 (seri numaralı, düzen /PANEL) + MB52 (toplu + değer, düzen PANEL / "MLZ STOK PANEL").
  MB52'de seri gruplar M03550000 ve M03250000 düşülür; M03550100 MB52'de kalır.
  Değer: depo+malzeme+parti eşleşmesi (YENİ / ÇIKMA-BDSZ / HURDA karışmaz), sonra boş parti, sonra il geneli.
  MB52 serbest stokta olmayan seri: satır görünür, adet 0, tutar 0. Sonuçta mutabakat: AKSORA vs MB52.
- Malzeme: CDP 7151–7155, EPP 7051–7055. Bütün iller: sistem başına tek MB52 (5 il). Tek il düğmeleri:
  sistem başına tek MB52 (o il), CDP/EPP normal-ilçe-yüklenici bellekten dilimlenir (AKHZ_TEK_AL).
- Güncelleme bitişi: panel düğmeleri → ana panel + tek özet; stok sayfası düğmeleri → aynı sayfa +
  tek özet; bütün iller → ana panel. Süre şeritte görünür.
- İl seçimi: seçim ekranı tek görünür sayfa; il tıklaması ekran/olay kapalı tek seferde.
- SAP güvenlik penceresi otomatik tıklanmaz (kural: kullanıcı SAP GUI kuralı ekler: C:/TEMP/* oku-yaz izin).

## Müşteri birimi il bilgileri
| İl | Tesis | CDP depoları | İlçe | Yüklenici |
|---|---|---|---|---|
| Artvin | 7155 | E###, D003, D005, S001, Z002, Z003, Z007, EM14 | K### (yalnız M03550000/M03550100/M03250000) | M132 |
| Trabzon | 7152 | E###, D003, D005, D006, D012, P001–P003, S001, Z002, Z003 | K### (aynı 3 mal grubu) | M090, M132, M501–M511 |

- Motor il bazlıdır: tesis `MS_Tesis()`, depo eşlemesi `MS_Hedef` (il'e göre), desteklenen iller
  `AKSORA_MUSTERI_IL_DESTEK` / `AKSORA_MUSTERI_ILLER`. Yeni il = bu üç yere ekleme.
- Her ilin stoğu kendi gizli kaydında: `_MA_<IL>_CDP / _ILCE / _YUK`. İl değişiminde
  (`AKSORA_MUSTERI_IL_DEGIS`) ekrandaki il kaydedilir, seçilen il yüklenir; M_ sayfalarında BX20 = il.
- Müşteri "bütün iller": tanımlı illeri müşteri motoruyla sırayla günceller, sonunda seçili ile döner ve müşteri ANA PANEL'i açar.
- Stok sayfalarında C2 = C2:E2 birleşik: temizlik daima `C2.MergeArea`. Arşiv kaydet/yükle hataları yutulmaz; yükleme önce 3 kaydı okur sonra 3 sayfayı yazar; il seçimi (AZ1) ancak 3 sayfanın BX20'si doğrulanınca değişir.
- Müşteri arama sayfası: il tarihleri ekrandaki il için sayfalardan, diğer iller için `_MA_<IL>_*` A1'den. Toplu düğme = malzeme arama düğmesinin aynısı.
- Depo listesi (form açılır liste) tüm satırları açar (tekerlek kaydırmaz); sayfa her açılışta yeniden yerleştirilir.
