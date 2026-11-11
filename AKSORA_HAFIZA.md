# AKSORA – Sistem Hafızası (kilitli temel, 13.10.2026)

Bu dosya, AKSORA'nın kilitlenmiş tasarım ve çalışma kurallarıdır. Yeni il / yeni birim
eklenirken bu kurallar birebir uygulanır; "aynı olanlara dokunulmaz".

## Kullanıcı ve çalışma şekli
- Burhan Serdar, Çoruh EDAŞ malzeme stok uzmanı. Türkçe, kısa, açıklamasız; "ne yapacağım" net söylenir.
- VBA proje şifresi: kullanıcıda (yalnız VBE kilidi için).
- Temel dosya: kullanıcının en son yüklediği `AKSORA.xlsm` (5 ilin müşteri kayıtları içinde;
  diğer iller tekrar güncellenmez). Repo `final/AKSORA.xlsm` = bu dosya; yeni sürümler bunun üzerine kurulur.
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
- Panel şeridi: "İl güncelleme süresi" + "Bütün iller" (toplu süre ayrı Name: _AKSORA_MALZEME_TOPLU_SURE / _AKSORA_MUSTERI_SON_SURE).
- Güncelleme bitişi: panel düğmeleri → ana panel + tek özet; stok sayfası düğmeleri → aynı sayfa +
  tek özet; bütün iller → ana panel. Süre şeritte görünür.
- İl seçimi: seçim ekranı tek görünür sayfa; il tıklaması ekran/olay kapalı tek seferde.
- SAP güvenlik penceresi otomatik tıklanmaz (kural: kullanıcı SAP GUI kuralı ekler: C:/TEMP/* oku-yaz izin).

## MÜŞTERİ BİRİMİ — KİLİT (13.10.2026, kullanıcı onaylı)
- 5 il (Artvin 7155, Trabzon 7152, Gümüşhane 7153, Giresun 7154, Rize 7151) aktif ve MB52 ile mutabık.
- Ana panel, stok sayfaları, raporlar (alt bölüm: depo bazında malzeme dağılımı), arama, il seçimi,
  bütün iller güncellemesi bu haliyle kilitli. Kullanıcı istemedikçe müşteri birimine dokunulmaz.
- Kilit sürümü: commit dc1d8e4 (final/AKSORA.xlsm). Yeni işler bu dosyanın / kullanıcının son yüklediği dosyanın üzerine kurulur.

- Birim seçim ekranında Müşteri Birimi AKTİF (BIRIM_CIZ sürüm BX2=8).

## Malzeme birimi depo kuralı
- İlçe tarafı (CDP ve EPP): K001–K999 + PB01–PB999 (AKSORA_IL_DEPO_GECERLI "ILCE"). Hızlı okuma (AKHZ) bu kuralı kullanır;
  eski yedek yöntem SAP'de yalnız K aralığını sorgular.

## Müşteri birimi il bilgileri
| İl | Tesis | CDP depoları | İlçe | Yüklenici |
|---|---|---|---|---|
| Artvin | 7155 | E###, D003, D005, S001, Z002, Z003, Z007, EM14 | K### (yalnız M03550000/M03550100/M03250000) | M132 |
| Trabzon (onaylı) | 7152 | E###, D003, D005, D006, D012, P001–P003, S001, Z002, Z003 | K### (aynı 3 mal grubu) | M090, M132, M501–M511 |
| Gümüşhane (onaylı) | 7153 | E###, D003, D005, D008, D015, Z003 | K### (aynı 3 mal grubu) | M132, M501, M502 |
| Giresun (onaylı) | 7154 | E###, D003, D005, D008, D012, P001, Z003, Z007 | K### (aynı 3 mal grubu) | M031, M132, M501–M505 |
| Rize (onaylı) | 7151 | E###, D003, D005, D012, D015, Z002, Z003 | K### (aynı 3 mal grubu) | M010, M036, M038, M040, M132, M501–M504 |

- Motor il bazlıdır: tesis `MS_Tesis()`, depo eşlemesi `MS_Hedef` (il'e göre), desteklenen iller
  `AKSORA_MUSTERI_IL_DESTEK` / `AKSORA_MUSTERI_ILLER`. Yeni il = bu üç yere ekleme.
- Her ilin stoğu kendi gizli kaydında: `_MA_<IL>_CDP / _ILCE / _YUK`. İl değişiminde
  (`AKSORA_MUSTERI_IL_DEGIS`) ekrandaki il kaydedilir, seçilen il yüklenir; M_ sayfalarında BX20 = il.
- Müşteri "bütün iller": tanımlı illeri müşteri motoruyla sırayla günceller, sonunda seçili ile döner ve müşteri ANA PANEL'i açar.
- Stok sayfalarında C2 = C2:E2 birleşik: temizlik daima `C2.MergeArea`. Arşiv kaydet/yükle hataları yutulmaz; yükleme önce 3 kaydı okur sonra 3 sayfayı yazar; il seçimi (AZ1) ancak 3 sayfanın BX20'si doğrulanınca değişir.
- Müşteri: MB52 seri grubunda IQ09 seri sayısını aşan miktar "seri no boş" satır olarak yazılır (adet+tutar MB52'den) → mutabakat birebir.
- Müşteri arama: tüm tanımlı iller aranır (ekrandaki il sayfalardan, diğerleri arşivden); il düğmesi BX21'de saklanır, sonuç o ile süzülür; sonuç yalnız başka sayfadan gelince temizlenir.
- Müşteri arama sayfası: il tarihleri ekrandaki il için sayfalardan, diğer iller için `_MA_<IL>_*` A1'den. Toplu düğme = malzeme arama düğmesinin aynısı.
- Depo listesi (form açılır liste) tüm satırları açar (tekerlek kaydırmaz); sayfa her açılışta yeniden yerleştirilir.

## Ek kurallar (13.10.2026)
- Ana panel orta sıra: SAP KONTROL = 1. kart, TÜM STOK = 2.(–3.) kart, MALZEME ARAMA = son kart hizasında (AKSORA_PANEL_ORTA_HIZALA).
- Rapor lejant kareleri 10x10, ad satırına ortalı (M_LEJANT_KARE).
- İl seçim ekranı tarihi: malzeme AKSORA_IL_SON_TARIH (en son güncellenen grup), müşteri AKSORA_MUSTERI_IL_SON_TARIH.
- Müşteri stok filtresi sonrası 2 sn içinde kendiliğinden arama sayfasına geçiş engellenir (AKSORA_MUSTERI_GECIS_KORU).

- Müşteri raporları alt bölüm: DEPO BAZINDA MALZEME DAĞILIMI (M_DEPO_DAGILIM). Depo listesi (AKSORA_RP_DM_LISTE) →
  malzeme pivot tablosu (Malzeme, Tanım, Miktar, Pay, Toplam Değer + TOPLAM) + sağda halka grafik (ilk 9 + Diğer) ve lejant.
  Seçim BX30, bölüm satırları BX31:BX32. Şekiller AKSORA_RP_DM_*; üst halka döngüsü bunlara dokunmaz.

## Yatırım birimi (15.10.2026)
- Giriş: YATIRIM IL SECIM → il → PANELE GİRİŞ → Y_GIRIS (Y_KUR, Y_IL_UYGULA, Y_IL_VERI, panel, arama tarihleri, Y_ALT_SAYFALAR).
- Sekme sırası: ANA PANEL, TEMİNAT, HAKEDİŞ, MALZEME ARAMA, CDP stok/rapor, EPP stok/rapor, CDP yük. stok/rapor, EPP yük. stok/rapor.
- Depolar: CDP ve EPP yatırım = D002, D011; yüklenici = malzeme birimindeki ilin yüklenici depoları (AKSORA_IL_YUKLENICI).
- Veri: MB52 PANEL düzeni (AKSORA_YATIRIM_MB52) → 14 sütun müşteri stok düzeni; il kayıtları _YA_<IL>_<1..4> (paketli).
- Arama: Y_MALZEME_ARA yalnız yatırım verisini tarar; D depoları her zaman üstte.
- Teminat Takibi (Y_TEMİNAT, TM_V3): Z ile başlayan depolar (Z007 vb.) yatırım/yüklenici kapsamında DEĞİL; stoksuz depolar listelenmez. HER DEPO AYRI (ayrı yüklenici/hakediş; gruplama YOK). (1) Depo bazında teminat: teminat (giriş),
  DEPODA (teminata tabi), SAHADA (hakediş bekleyen), kalan, kullanım % (depoda ÷ teminat), durum (GÜVENLİ <%80, DİKKAT %80-100, AŞIM >%100),
  sahada >45 gün sayısı. (2) Malzeme detayı sütun sırası: Depo, Depo Adı, Malzeme, Tanım, Miktar (birimiyle), Tutar, Konum
  (Depoda/Sahada/Kısmi), Sahadaki Miktar, Depodaki Tutar, Sahadaki Tutar, Sahada (gün); depo seçici filtre.
  Kayıt _Y_TEMINAT: T|il|DEPO (teminat), K|il|DEPO|MALZEME (konum, sahadaki miktar, işaret tarihi).
  Miktarlar: tam sayıda ondalık ayıraç gösterilmez (Y_MIKTAR_BICIM). AKSORA amblemi içeriğin altına (Y_FILIGRAN).
- Hakediş (Y_HAKEDİŞ): şimdilik PROJE – STOK KARŞILAŞTIRMASI. PROJE EXCELİ YÜKLE: kod/tanım/miktar/birim/fiyat sütunları otomatik bulunur
  (başlıksız listede 7-12 haneli kod sütunu). Karşılaştırma: seçili yüklenici deposu (veya tümü) + yatırım depoları (D002/D011);
  durum EKSİK / Yatırımdan tamamlanır / Depoda yeterli / Projede yok (iade). Proje _Y_PROJE'de il bazında. Hakediş hesabı uzman görüşü sonrası
  (sayfada akış kartları ile yer ayrıldı).
- Teminat girişleri Workbook_SheetChange → AKSORA_YATIRIM_DEGISIM ile anında kaydedilir.

## Dosya boyutu ve kayıt biçimi (15.10.2026)
- İl kayıtları (_MA_*, _YA_*) PAKET1 biçiminde: A1 tarih, B1 il, C1 "PAKET1", D1 satır sayısı; A2.. bloklar
  (satır = 14 alan TAB ile, alan = tür harfi s/n/b + değer, n nokta ondalık; blok ≤30.000 karakter, satırlar LF).
  Okuma/yazma yalnız AKSORA_PAKET_OKU / AKSORA_PAKET_YAZ / AKSORA_PAKET_ADET ile. Eski biçim açılışta AKSORA_ARSIV_BAKIM ile çevrilir.
- Base dosya tools/optimize_base.py ile hazırlanır (paketleme + Y stok satır kırpma + görsel küçültme/JPEG).
- Bant görselleri .jpg varsa o kullanılır (AKSORA_GORSEL). Sonuç: 32,5 MB → 11,4 MB.
- Masaüstü başlatıcı: AKSORA_BASLAT.hta (+ AKSORA_LOGO.png, AKSORA.ico) Excel'i görünmez açar, AKSORA yükleme ekranı gelince kapanır.

## İlçe Yönetim birimi (15.10.2026)
- Hub'da 4. kart (04 İlçe Yönetim Birimi, AKSORA_BIRIM_ILCE); hub BX2 sürümü 9 (4 kart, genişlik (1080-68-54)/4).
- İl seçimi: "ILCE IL SECIM" (KAPAK kopyası, YATIRIM_KAPAK_AC ilce:=True); bant AKSORA_ILCE_BANT.png varsa eklenir.
- Sayfalar (I_, etiket AKSORA_ILCE_V1, sıra sabit): I_ANA PANEL, I_MALZEME_ARA, I_STOK_KARŞILAŞTIRMA, I_CDP_İLÇELER_STOK,
  I_CDP_İLÇELER_RAPOR, I_EPP_İLÇELER_STOK, I_EPP_İLÇELER_RAPOR, I_İSG_MALZEMELERİ (CDP ve EPP AYRI sayfalar; kayıt tek _IA_<IL>).
  Panel (IP_V3): büyük düğmeler STOK / DEMONTAJ KONTROLÜ, TÜM İLÇE STOKLARINI GÜNCELLE, MALZEME ARAMA, CDP/EPP İLÇE GÜNCELLE,
  İSG ZİMMET KONTROLÜ; küçük hızlı erişim yalnız CDP/EPP İLÇE STOK-RAPOR (4 düğme, büyük düğmesi olanlar tekrarlanmaz).
  Q kartı başlığı İSG ZİMMET KONTROLÜ. Hub 04 kartı: İlçe depo stokları / İlçeler arası malzeme arama / Demontaj malzeme kontrolü /
  İSG zimmet kontrolü (hub BX2 sürüm 10). İl seçim bandı AKSORA_ILCE_BANT.jpg (KAPAK fotoğrafından, üst şerit: İlçe Stok Takibi,
  İlçeler Arası Malzeme Arama, Demontaj Malzeme Kontrolü, İSG Zimmet Kontrolü, Analiz ve Raporlama; tools/kapak_ilce.py).
  AKSORA amblemi (FILIGRAN) ilçe sayfalarında kullanılmaz. Depo sayısı = ilin bilinen tüm ilçe depoları (_I_DEPO kalıcı liste; stoksuz depo da sayılır).
  Yatırım ve ilçe aramaları her zaman il kaydından okur (stok sayfasındaki filtre aramayı daraltmaz). Panel/arama/stok/rapor müşteri şablonundan kopya (Y_KUR ilce:=True).
- Veri: CDP 7151-7155 + EPP 7051-7055, yalnız K001-K999 depoları; ikisi tek stok sayfasında (depo|tesis|malzeme sıralı).
  Kayıt _IA_<IL> (PAKET1). Tek sistem güncellenince diğerinin son kaydı korunur.
- Günlük geçmiş _IH_<IL>: her sütun bir gün (günün son güncellemesi), en fazla 20 gün; satır depo TAB tesis TAB malzeme TAB
  depo adı TAB tanım TAB birim TAB miktar TAB değer.
- Stok karşılaştırma (IK_V3): tablo A:N tek parça (bantlar N dahil), ANA PANEL sağda, ilçe seçimi tek açılır liste (I_KV_SECIM, O8). Malzeme biriminin STOK_KARŞILAŞTIRMA sayfasının kopyası (birebir görünüm). İl düğmeleri yerine
  ilçe depo düğmeleri (O sütunu, I_KV_D_*). Tarihler B6/I6 (liste X sütunu, en yeni üstte; varsayılan yeni = son, eski = önceki günün son).
  Durum: AZALDI / ARTTI / DEMONTAJ GİRİŞ. N sütunu DEMONTAJ KONTROLÜ: azalan (çıkma olmayan) her malzeme için aynı depo+tesis+malzemede
  ÇIKMA BEDELSİZ (SAP: CIKMA/BDSZ) değerlemeli artış aranır; yoksa kırmızı "DEMONTAJ GİRİŞİ YAPINIZ • adet", az ise "EKSİK DEMONTAJ • giren/azalan".
  Filtreler: TÜMÜ, AZALAN, ARTAN, DEMONTAJ GİRİŞ, DEMONTAJ YAPINIZ. Geçmiş kaydı değerlemeyi de tutar (9 alan).
- İSG (I_İSG_MALZEMELERİ, 15.10.2026): ana liste docs/isg/ISG_MALZEME_KOD.xlsx/.csv (293 malzeme, 12 mal grubu).
  KİŞİSEL ZİMMET listesi docs/isg/ZIMMET_GEREKEN.xlsx/.csv (kullanıcının dosyası, 2302 kod, "ZEAM SARFI").
  Birleşik liste docs/isg/ISG_ISLEM.csv: 293 İSG malzemesinin 261'i ZİMMET, 32'si MANUEL SARF. VBA'da I_ISG_LISTE (kod+Z/M gömülü).
  Sayfa: KPI (İSG kalem, değer, zimmet yapılacak, manuel sarf yapılacak); İLÇE BAZINDA İSG DURUMU özeti (zimmet/sarf kalem-miktar,
  değer, durum ZİMMET BEKLİYOR / SARF BEKLİYOR / TEMİZ); İSG MALZEME LİSTESİ (ilçe açılır liste + TÜMÜ / ZİMMET YAPIN / MANUEL SARF YAPIN
  filtreleri; satırda YAPILACAK İŞLEM: kırmızı ZİMMET YAPIN, mavi MANUEL SARF YAPIN). Veri: _IA_<IL> kaydı (CDP+EPP), kod Y_KOD_NORM ile.
  Liste değişirse docs/isg güncellenip I_ISG_LISTE yeniden üretilir (fx/isg_gen.py mantığı).
- Panel bandı AKSORA_ILCE_PANEL_BANTLARI.jpg (yoksa müşteri bandı).

## Canlı güncelleme (15.10.2026, tüm birimler)
- Her birimde güncelleme (tek il ve bütün iller) ANA PANEL'de izlenir: panel ekrana gelir, kartlar mavi LED + "Güncelleniyor...",
  şerit müşteri paneli teması (açık sarı RGB(255,246,219), kahverengi kalın yazı RGB(146,64,14); tamam yeşil (209,250,229)/(6,95,70);
  hata (254,226,226)/(153,27,27); bilgi (224,242,254)/(12,45,79)). Metin "Güncelleniyor  •  [Bütün iller n/5  •]  il  •  adım".
- Malzeme: AKSORA_PANEL_GUNCELLENIYOR + AKSORA_PANEL_SERIT (B6); bütün iller artık MALZEME_ARA yerine ANA PANEL'de.
  Müşteri: AKSORA_MUSTERI_PANEL_ON (MS_CEKIRDEK başı, bütün iller başı). Yatırım/İlçe: Y_CANLI_BASLAT, Y_DURUM/I_DURUM, mYSira.
- İlçe paneli büyük düğmeler (BX27 DUGME_V1): üst sıra STOK / DEMONTAJ KONTROLÜ | MALZEME ARAMA | İSG ZİMMET KONTROLÜ (aynı beyaz),
  alt sıra CDP İLÇE GÜNCELLE | TÜM İLÇE STOKLARINI GÜNCELLE (yeşil) | EPP İLÇE GÜNCELLE.
- Malzeme il bandı: AKSORA_PANEL_BANNER_<IL>.png yerine birleşik AKSORA_MALZEME_PANEL_BANTLARI.jpg (AKSORA_MALZEME_BANT_EKLE).

## Klasör düzeni (15.10.2026, paylaşım için)
AKSORA\  AKSORA.xlsm, AKSORA_BASLAT.hta
AKSORA\GORSELLER\BANNER  AKSORA_BIRIM_BANNER_03, AKSORA_PANEL_TEK_PARCA_V7, AKSORA_PANEL_SADE_P1, AKSORA_RAPOR_BANNER,
                         AKSORA_MALZEME/MUSTERI/YATIRIM/ILCE_PANEL_BANTLARI
AKSORA\GORSELLER\KAPAK   AKSORA_MUSTERI_BANT, AKSORA_YATIRIM_BANT, AKSORA_ILCE_BANT
AKSORA\GORSELLER\LOGO    AKSORA.ico, AKSORA_LOGO.png, AKSORA_SPLASH_LOGO.bmp, AKSORA_RAPOR_FILIGRAN.png
AKSORA\KAYITLAR           AKSORA_SAP_TESHIS.txt, AKSORA_BIRIM/TASARIM_ONCESI_* yedekleri (klasör varsa oraya yazılır)
- Kod tüm görselleri AKSORA_GORSEL(ad) ile arar: kök, GORSELLER\BANNER, \KAPAK, \LOGO, GORSELLER (jpg tercih). Yeni görsel
  yolu yazarken ThisWorkbook.Path & "\..." KULLANMA, AKSORA_GORSEL kullan. Kayıt klasörü: AKSORA_KAYIT_KLASORU().
- HTA: ICON ve logo GORSELLER\LOGO (logo kökte de aranır); masaüstü kısayol simgesi yeni yere bağlanır.
- final/duzen/AKSORA_KLASOR_DUZENLE.vbs: kullanıcının klasörünü bu düzene getirir (silmez; eski dosyalar _SILINEBILIR'e).
  Gereksiz: AKSORA.vbs, KUR.vbs, AKSORA_MASAUSTU.ico, AKSORA_PANEL_BANNER.png, AKSORA_PANEL_BANNER_<5 il>.png.

## 16.10.2026 düzeltmeleri (kullanıcının yüklediği dosya yeni baz)
- Baz: kullanıcının 16.10 AKSORA.xlsm'i (Sayfa85-93 eklendi). optimize_base: hub PNG'ler JPEG, Y_/I_ *_STOK fazla satırları;
  ZATEN PAKETLİ (_MA_/_YA_ C1="PAKET…") arşivlere dokunmaz (aksi halde boşaltıyordu!). 18,4 MB -> 13,2 MB.
- Y_/I_ RAPOR "DEPO BAZINDA MALZEME DAĞILIMI" kayması: AKSORA_YATIRIM_DM_DUZEN artık tüm DM şekillerini satırdan
  (BX31 baş, BX32 son, HALKA_ZEMIN sol/genişlik) yeniden hesaplar: zemin=baş satır üstü, halka=+76, lejant adımı min(24,(gs-20)/m).
  Y_DM_KONTROL (AKSORA_YATIRIM_GORUNUM içinde) sayfa açılınca kayma varsa sessizce toparlar.
- Malzeme ANA PANEL güncellenirken kart: "Güncelleniyor..." RGB(29,78,216), LED (59,130,246)+glow (müşteri ile aynı).
  AKSORA_PANEL_GUNCELLENEN listesi + AKSORA_PANEL_KART_CANLI(kod); PANEL_KART_GUNCELLE, AKSORA_DURUM_RENK_GUVENLI,
  AKSORA_ANA_PANEL_DURUM_RENK_YENILE "1 GÜN ÖNCE" yazmaz; CANLI iken B6 şeridi ezilmez.
- I_STOK_KARŞILAŞTIRMA tablo IK_V4 (I_KARS_SUTUN): DURUM | ÜY | DEPO | İLÇE | MALZEME | AÇIKLAMA | DEĞERLEME | ESKİ/YENİ MİKTAR | FARK |
  ESKİ/YENİ DEĞER | DEĞER FARKI | DEMONTAJ KONTROLÜ. BİRİM sütunu yok: birim miktar biçiminde ("12 ADT"). Filtre alanları 1 ve 14 aynı.
- Malzeme stok sayfaları ÇIKTI AL (btnStokCikti, T_StokYap; AKSORA_STOK_CIKTI, modAKSORA_TASARIM): kullanıcının SAP sayım listesi
  standardı: Üretim yeri | Depo yeri tanımı | Malzeme | Malzeme kısa metni | Parti | Tahditsiz klnb. (miktar+birim tek sütun) |
  SAYILAN | FARK (geniş, elle yazılır). HURDA ve CIKMA/BDSZ partiler alınmaz; YENI ve boş parti alınır. Ekrandaki (depo/filtre)
  satırlar. Renksiz, gri başlık, tüm kenarlıklar, A4 yatay, tek sayfa genişliği, başlık satırı her sayfada.
  PrintPreview KULLANMA: gizli şeritte (UYGULAMA_MODU) Excel donar. Gizli geçici kitap -> PDF (TEMP) -> OpenAfterPublish.
- ÇIKTI AL tüm sayfalarda (16.10): AKSORA_CIKTI_DUGME (btnCiktiAl; sayfa açılışında BIRIM_OLAY/EKRAN_SIGDIR'dan, ANA PANEL ve
  yanındaki düğmelerin solunda; karşılaştırmada ANA PANEL altında; malzeme stokta btnStokCikti). AKSORA_CIKTI_AL türe göre:
  1 stok (malzeme/M_/Y_/I_) -> sayım listesi (M_ seri no varsa "Seri numarası" sütunu), 2 arama/karşılaştırma -> ekrandaki tablo,
  3 rapor/İSG/teminat/hakediş -> sayfanın kendisi (düğmeler PrintObject=False). Panellerde yok.
  ONAYLI DÜZEN = ilk tasarım (imza kutusu satır arası YOK; FitToPagesWide=1, BlackAndWhite=True, satır 24).
  Sayım sütunları: Üretim yeri | Depo (kod) | Depo yeri tanımı | Malzeme | Malzeme kısa metni | [Seri no] | Parti |
  Tahditsiz klnb. | SAYILAN | FARK. TOTAL İPTAL (kullanıcı istemedi). Sıra: malzeme kısa metni A-Z, malzeme, depo.
  Alt bilgi: orta = imza bandı RESMİ (&G; kodda base64 PNG -> TEMP\AKSORA_IMZA_BANDI_V3.png, 27,7x1,3 cm, şeffaf zemin, etiket 4,5pt (6pt bile büyük bulundu); dikey çizgiler
  3,2/13,85/25,6 cm, etiketler İLGİLİ DEPO SORUMLUSU/LARI | KONTROL EDEN SORUMLU/LAR), sol 7pt Çıktı, sağ 7pt Sayfa.
  Alt kenar 1,65 cm (ilk tasarım 1,4) -> sayfa sayısı ~aynı (kullanıcı 19 sayfa ister; 4,4 cm alt kenar 35 sayfa yaptı!).
  Çok satırlı/büyük punto alt bilgi denemesi orta/sağ bölümü hiç basmadı; kullanma. Kaynak resim docs/cikti/.
- Y_/I_ rapor üst halka lejantında yüzdeler (HALKA_PAY_n) PAY başlığından kayıyordu: Y_PAY_HIZALA (DM_DUZEN başında)
  ve Y_DM_KONTROL (sayfa açılışı) PAY_n'yi PAY_BASLIK sol/genişlik/sağa hizaya getirir.
- Çıktı başlığı "Depo:" = çıktıya giren depoların TANIMLARI (virgülle; sayfadaki tüm depolar ise TÜMÜ). H2 hücresine bakma
  (filtreyle seçilen depoyu göstermiyordu).

## ADMİN | Yönetim Merkezi (16.10.2026)
- Hub (BIRIM SECIM, BX2 sürüm 11): sağ altta "ADMİN GİRİŞİ" -> AKSORA_ADMIN_GIRIS, şifre adminbs1284 (InputBox; VBA'da maskeli kutu yok).
- Sayfa "ADMIN PANEL" (ilk girişte kurulur, BIRIM_DIGERLERINI_GIZLE listesinde). Kod modAKSORA_BIRIM sonunda (ADM_*).
  İçerik: başlık + düğmeler (TÜM BİRİMLERİ GÜNCELLE / RAPORU YENİLE / ÇIKTI AL / BİRİM SEÇİMİNE DÖN), otomatik yönetici özeti,
  sağlık skoru (100 - 12*kritik - 4*uyarı), 4 birim kartı (il, güncellik, değer, satır/depo, birime özel risk),
  arama (C20, Enter veya ARA; tüm birimlerin ekrandaki stok sayfalarında kod/ad/seri; birim+depo+malzeme gruplu ilk 20),
  uyarılar (güncellik, eksi stok, müşteri fiyatsız, malzeme HURDA, İSG zimmet/sarf, demontaj eksik, teminat aşım/dikkat/45 gün,
  karşılaştırma sıfırlanan), İSG ve demontaj tabloları, en değerli 10 depo (veri çubuklu).
- Tek tuş: AKSORA_ADMIN_TOPLU=True iken 4 birimin bütün iller makroları onay/özet MsgBox göstermez, sonuç AKSORA_ADMIN_SONUC'a yazılır.
- I_İSG notu (Y_BOLUM notSutun=9) tablonun son sütununa yaslı.
- Baz: kullanıcının 16.10 öğlen dosyası (up18), bx/xl = u18/xl, frmSplash VB_Base u18'den.

### ADMİN v2 (19.10.2026)
- İl filtresi C4 (TÜMÜ + 5 il; liste BX10:BX15; seçim BX3; SheetChange -> AKSORA_ADMIN_DEGISIM). Arama kutusu C30, sonuç 31-52.
- Veri: ADM_VERI(u, il) — birimin ekrandaki ili (stok sayfası A5) canlı sayfadan, diğer iller kayıtlardan:
  malzeme _AKLIVE_<IL>_{CDP,EPP,ILC,YUK} (stok düzeni, tarih A3), yatırım _YA_<IL>_1..4, müşteri _MA_<IL>_{CDP,ILCE,YUK},
  ilçe _IA_<IL> (PAKET, tarih A1). mAdmOnbellek: imza (tarih+satır) değişmedikçe bellekte; satırlar tab-ayrık metin.
- Tekil değer: il+ÜY+depo birden çok birimde ise bir kez (en yüksek değerli kaynak). Birim kartları birimin kendi görünümü.
- Bölümler: özet, skor (100-10*kritik-3*uyarı), birim kartları, İL KARŞILAŞTIRMASI matrisi, arama, uyarılar,
  İSG (TÜMÜ'de il bazında, il seçiliyken ilçe bazında; _IA_ + I_ISG_LISTE ile hesaplanır), demontaj (karşılaştırma ekranındaki il),
  en değerli 10 depo (tekil), en değerli 10 malzeme (tekil depolardan).
- Baz: kullanıcının 19.10 dosyası (up19).

### ADMİN v3 (19.10.2026)
- İl seçimi açılır liste DEĞİL: satır 4'te il düğmeleri (ADM_IL_n, AlternativeText = il kodu, AKSORA_ADMIN_IL_SEC). (Doğrulama listesi
  ilk değişiklikten sonra kayboluyordu.)
- Sıra: arama en üstte (C7, sonuç 8-39, 30 grup; çok kelime = hepsi; il dağılımı; TEMİZLE; arama BX4'te saklanır, il değişince tekrar),
  sonra DEMONTAJ KONTROLÜ (ADM_DEMONTAJ: her il _IH_ geçmişinden, son güncelleme vs önceki günün son kaydı, I_KARS_KUR kuralı;
  TÜMÜ'de il satırları, il seçiliyken ilçe deposu satırları + girişi yapılmayan ilk 15 malzeme), özet+skor, birim kartları,
  il karşılaştırması (her zaman 5 il), uyarılar (İSG/demontaj il başına TEK uyarı), İSG tablosu + en çok zimmet bekleyen 5 depo,
  en değerli 10 depo/malzeme.
- Skor: kategori başına bir kez (kritik -15, uyarı -6).

### ADMİN v4 (19.10.2026)
- Sayılar: miktar hücreleri ADM_MIKTAR (AKSORA_V13_MIKTAR_BICIMI; tam sayıda sonda virgül yok). "#,##0.###" hücre biçimi KULLANMA.
- Arama: birimler arası aynı il+ÜY+depo+malzeme TEK satır ("MALZEME + YATIRIM", miktar bir kez); sıralama ANA DEPO (D, mavi) ->
  İLÇE (K, mor) -> YÜKLENİCİ (diğer, sarı), her grupta miktar büyükten; "Depo türü" sütunu; birleşik hücre yok (G sütunu 30 geniş);
  başlıkta AutoFilter (koruma AllowFiltering, EnableSelection serbest); en fazla 398 satır (ADM_ARA_N 400 gizli yedek satır).
- Demontaj: eksik değer (eski değer/miktar birim fiyatı x eksik), "en çok" sıralı, eksik listesi en yüksek değerli 25.

- DEMONTAJ KONTROLÜNDE İSG MALZEMELERİ HARİÇ (I_ISG_LISTE): hem admin ADM_DEMONTAJ hem İlçe STOK/DEMONTAJ ekranı (I_KARS_KUR:
  "İSG malzemesi (demontaj gerekmez)", gri, sayılmaz). Admin aramada miktarı 0 olan satırlar gösterilmez; büyük tutarlar ShrinkToFit.

- Admin güvenlik (19.10): mAdmOturum (şifreyle True, ÇIKIŞ'ta False); şifresiz sekme/Ctrl+PgDn geçişi BIRIM_OLAY'da hub'a döner.
  Giriş kaydı _AKSORA_ADMIN_LOG (son 15: tarih|kullanıcı|GİRİŞ/HATALI); hatalı denemeler uyarıda, son 5 giriş sayfa altında.
  Demontaj listesinde EKSİK GİRİŞ turuncu, GİRİŞ YOK kırmızı.

- Admin şifre kutusu (19.10, düzeltme): ActiveX panel kullanıcıda çalışmadı, kaldırıldı. Şimdi InputBox
  (başlık "AKSORA - YONETICI GIRISI") + SetTimer/AddressOf ADM_MASKE_ZAMAN: kutu açılınca Edit'e EM_SETPASSWORDCHAR '*'.
  "Şifreyi göster" onay kutusu artık şifre kutusunun İÇİNDE (CreateWindowExW "Button" BS_AUTOCHECKBOX, kutu
  edit yüksekliği kadar uzatılır; zamanlayıcı 40ms BM_GETCHECK okur, değişince EM_SETPASSWORDCHAR 0/'*').
  Her girişte gizli başlar. Hub'daki göster düğmesi kaldırıldı (hub sürümü 13). Declare'ler yalnızca VBA7.
- Admin bekleme ekranı (19.10): ayrı pencere (frmSplash) kullanıcıda başka monitörde açıldı, istenmedi. Şimdi aktif sayfanın
  görünen alanına ADMY_* şekilleri (perde + kart + durum + çubuk, zoom'a göre ölçekli). ADM_ILERLE ScreenUpdating=True+DoEvents ile
  ilerletir (okuma %5-65, demontaj %66-82); %88'de ScreenUpdating=False, ekran kartta donarken rapor çizilir.
  ADM_YUKLEME_KAPAT kalan ADMY_ şekillerini siler ve korumayı geri koyar. Interactive=False, mAdmMesgul yeniden girişi engeller.
- Genel bekleme kartı (19.10): AKSORA_BEKLE_AC(başlık, mesaj, y0) / AKSORA_BEKLE_ILERLE / AKSORA_BEKLE_KAPAT (modAKSORA_BIRIM).
  Kart çizildikten sonra RedrawWindow(Application.hWnd, RDW_INVALIDATE|ERASE|ALLCHILDREN|UPDATENOW) ile hemen basılır
  (önbellekli ikinci admin girişinde kart görünmüyordu). Koruma ayarları saklanıp UserInterfaceOnly ile geri konur.
  Sarmalanan girişler (asıl kod *_ASIL): AKSORA_V8_GIRIS, AKSORA_MUSTERI/YATIRIM/ILCE_GIRIS, AKSORA_BIRIM_MALZEME/YATIRIM/
  MUSTERI/ILCE, AKSORA_ILCE_KARS_LISTE/SEC/YENILE. Toplu güncelleme / açılış / V11_MESGUL sırasında kart açılmaz.
- İlçe STOK KARŞILAŞTIRMA listesi: önce İLÇE A-Z, sonra durum (demontaj eksikleri önce) ve değer farkı.
- Hub (19.10): Yatırım ve İlçe Yönetim kartları AKTİF rozetli (hub sürümü 14). Paylaşım notu: SAP düzeni "PANEL"
  kullanıcıya özelse başka kullanıcıda yoktur; global "/PANEL" olarak kaydedilmeli (Müşteri zaten "/PANEL" ister).
- Güvenlik/sahiplik (19.10):
  - İmza + kendini onaran bekçi (modAKSORA_KORUMA): AKSORA_SAHIBI="Burhan SERDAR" sabiti; AKSORA_IMZA_DOGRULA
    açılışta (Workbook_Open, AKSORA_HAFIF_KORUMALARI_YENILE sonrası) çağrılır; gizli adlar _AKSORA_SAHIP/_AKSORA_IMZA,
    kitap Yazar/Şirket, KAPAK ZZ2000 (beyaz) yeniden yazılır → silinirse sahip adına döner.
  - İki çıktı: final/AKSORA.xlsm (ana, okunur kaynak) ve final/AKSORA_DAGITIM.xlsm (paylaşım). Dağıtım kopyası
    fx/strip.py ile tam satır + tırnaksız satır sonu açıklamaları ve boş satırlar silinmiş rawdist/'ten derlenir
    (1467 açıklama satırı → 0). Excel kaynağı yeniden derler (build_bin _VBA_PROJECT stub + offset sıfır).
  - docs/LISANS/AKSORA_SAHIPLIK.md: tarih + SHA-256 parmak izleri (ana ve dağıtım).
  - VBA şifresi (bserdar1284) ve "görüntüleme için kilit" Excel'de VBE'den elle uygulanır (ikili dosyada zorlanmadı:
    bozulma riski ve sınıflandırıcı engeli). Dağıtım kopyasına uygulanır, ana kopya okunur kalır.
- Admin bekleme kartı düzeltme (19.10): il değiştirmede ve ilk girişte kartın kaybolma sebebi ADM_KUR CIZIM
  bölümündeki şekil silme döngüsüydü (tüm şekilleri siliyordu, ADMY_ kartı dahil). Artık ADMY_ şekilleri atlanıyor;
  kart donma (ScreenUpdating=False ile rapor çizimi) boyunca ekranda kalır, en sonda AKSORA_BEKLE_KAPAT kaldırır.
  İlerleme çizim öncesi %90 "Rapor çiziliyor...".
- Admin bekleme kartı (19.10, 2. düzeltme): ikinci giriş/il değişiminde kart gelmeme sorunu. İki neden:
  (1) AKSORA_BEKLE_AC başındaki `If mAdmYukleme Then Exit Function` takılı durumda kartı engelliyordu →
  artık `If mAdmYukleme Then AKSORA_BEKLE_KAPAT` ile temizleyip devam ediyor; Son: da mAdmYukleme=False.
  (2) AKSORA_BEKLE_ILERLE'deki DoEvents, EnableEvents=False'a rağmen il butonunu re-entrant tetikliyordu →
  DoEvents kaldırıldı, sadece RedrawWindow (RDW_UPDATENOW|ALLCHILDREN|ERASE|INVALIDATE) ile anında boyanıyor.
- Admin bekleme kartı (19.10, 3. düzeltme): DoEvents kaldırınca hiç boyamadı → DoEvents asıl boyayıcıymış, geri alındı.
  AKSORA_BEKLE_ILERLE paint anında Interactive'i geçici True yapar (ADM dışı birimlerde zaten True). ADM_AC'den
  application.Interactive=False/True kaldırıldı (birim ekranlarıyla aynı davranış); kart boyamasını bastırıyordu.
- Admin bekleme kartı (19.10, 4. düzeltme, kök neden): Kart, veri önbellekteyken (ikinci giriş / il değişimi) çıkmıyordu;
  başka birimler gezilince canlı sayfa imzası değişip önbellek bozuluyor, okuma uzuyor ve çok sayıda DoEvents ile kart
  basılıyordu. Excel yeni şekli tek DoEvents ile basmayabiliyor. AKSORA_BEKLE_ILERLE ilk ve son (%85+) boyamada
  0,3 sn DoEvents döngüsüyle bekler (Interactive=False, yeniden girişi engeller).
- Yeni temel (19.10, a292e461): kullanıcının dosyası; VBA aynı (yalnız VBE harf düzeltmeleri), frmSplash VB_Base yeni GUID
  ({E102BC2F...}{EC4A7D4A...}) raw/frmSplash.txt'e alındı; bx/xl = u20/xl; base = optimize_base --img.
- Demontaj kontrol dışı mal grupları: M03550000 sayaç, M03250000 modem, M03550100 aparat/mühür (+ İSG). I_DEM_DISI(mal, isgL)
  günlük kontrol (I_KARS_KUR, gri metin), admin (ADM_DEMONTAJ, önbellek DEM4) ve aylık defterde kullanılır.
  Malzeme haritası _I_MALZEME (A mal, B mal grubu, C tanım, D birim); I_MAL_HARITA ilk açılışta _IA_ paketlerinden doldurur,
  her ilçe güncellemesinde I_MAL_DIZI ile güncellenir.
- Aylık demontaj defteri _ID_<IL> (I_DEFTER_GUNCELLE / I_DEFTER_GUN / I_DEFTER_OKU): her sütun bir gün; satır 1 gün, satır 2
  kaynak imza "eski|yeni", altı "depo TAB ÜY TAB mal TAB kullanılan TAB girilen TAB kullanılan TL". 400 gün tutulur.
  İmza değişen / yeni gün çiftleri yeniden hesaplanır (geçmiş 20 gün olduğu için ilk açılışta geriye doğru doldurur).
  Aylık eksik = Σkullanılan − Σgirilen (depo|ÜY|mal, ay içi geç girişler dahil). Uyum % = değer ağırlıklı.
- İlçe sayfası I_AYLIK_DEMONTAJ (I_ADLAR'a eklendi, M_CDP_RAPOR kopyası): KPI 4 kart, 12 ay düğmesi (AKSORA_ILCE_AY_SEC,
  BX30=yyyymm), AYLIK GİDİŞAT, İLÇE BAZINDA (A-Z), MALZEME BAZINDA, İLÇE × MALZEME DETAYI (AutoFilter). Uyum veri çubuğu.
- Admin: ADM_AYLIK_CIZ (DEMONTAJ GİRİŞİ YAPILMAYAN listesinden sonra): il/ilçe tablosu + önceki ay, 6 aylık gidişat,
  en çok eksik 10 ve en çok kullanılan 10 malzeme. Y_SAYFA_SIFIRLA artık ADMY_ bekleme kartını silmiyor.
- DERS (19.10): Sabit I_DEFTER_GUN ile Sub I_DEFTER_GUN aynı adı taşıdığı için modAKSORA_BIRIM derlenmedi ("Compile error in
  hidden module"). Sabit I_DEFTER_SURE yapıldı. Yeni kodda modül düzeyi adlarla (sabit/değişken/yordam) çakışma taranmalı.
- DERS 2 (19.10): I_DEFTER_GUN içindeki yerel `mE` VBA'da `Me` anahtar kelimesi sayılıyor → "Syntax error" (kullanıcı
  Debug > Compile ile buldu). qE/qY/qV yapıldı. Her yeni kodda ayrılmış kelime taraması (Me, To, Is, In, On, Or, Mod, Do,
  Name, Line, Input, Error, Date, Type...) ve modül düzeyi ad çakışması taraması yapılmalı. Kilitli projede Excel satır
  göstermez; kullanıcıdan Alt+F11 > Debug > Compile VBAProject ekran görüntüsü istenir.
- Kritik demontaj (19.10): I_KATEGORI(mal) → KABLO (mal grubu M0316…), DİREK (M0308…), PANO (M0327…), HÜCRE (M0314…),
  TRAFO (M0340…, akım/gerilim trafoları dahil); mal grubu yoksa tanım başı (KABLO/DIREK/PANO/HUCRE/TRAFO). Diğerleri (sigorta,
  ampul vb.) "DİĞER", ikinci planda/gri. I_AY_HESAP top/depo dizileri 12 alan (8-11 kritik TL/eksik TL/kalem/eksik kalem),
  I_UYUM_K kritik uyum, I_KAT_OZET kategori tablosu, I_OZET_DURUM (KRİTİK EKSİK / DİĞER EKSİK / TAMAM).
  Uygulanan yerler: AYLIK DEMONTAJ sayfası (KPI'lar kritik, KRİTİK MALZEMELER tablosu, sıralama kritik önce, Kategori sütunu)
  ve admin (il/ilçe tablosu kritik, KRİTİK MALZEMELER, gidişat kritik, ilk 15 kritik eksik/kullanılan, günlük eksik listesi
  kritik önce + durumda kategori). Kullanıcı isteğiyle günlük STOK/DEMONTAJ KONTROLÜ sayfasına dokunulmadı.
  Kontrol betikleri: fx/checks.py (ad çakışması + ayrılmış kelime), fx/checks2.py <yordamlar> (tanımsız ad), fx/checks3.py (blok).
- Aylık demontaj sadeleştirme (19.10): Kullanıcı "çok gereksiz bilgi" dedi. AYLIK DEMONTAJ (ilçe) artık yalnız kritik:
  KPI (kritik uyum %, kritik fire TL, ★ EN İYİ İLÇE, ▼ EN KÖTÜ İLÇE), ay düğmeleri, KRİTİK MALZEMELER (kategori: kullanılan /
  girilen / eksik(fire) m-adet, oran, fire TL, fireli ilçe), İLÇE KARNESİ (sıra: kritik uyum desc, eşitlikte fire asc; kablo m,
  direk/pano/hücre/trafo adet eksik, değerlendirme ÇOK İYİ/İYİ/ORTA/ZAYIF + EN İYİ/EN KÖTÜ), KRİTİK FİRE LİSTESİ (filtreli).
  Admin ADM_AYLIK_CIZ: il özeti (TÜMÜ), KRİTİK MALZEMELER, İLÇE KARNESİ (TÜMÜ ve >15 ilçe: en iyi 5 + en kötü 10).
  Gidişat, ilk-15 listeleri, malzeme/detay tabloları kaldırıldı (ADM_AY_SATIR / ADM_MAL_SIRA artık kullanılmıyor).
  Görsel: I_UYUM_CUBUK = trafik ışığı ikonu (70/90) + yeşil veri çubuğu; I_FIRE_ISI = beyaz→kırmızı renk skalası; kategori renkleri.
- Kaydetme yavaşlığı (19.10 analiz): dosya açık hali 94 MB XML; sharedStrings 28 MB (çoğu _MA_<IL>_CDP müşteri arşivleri
  ~19 MB: sayaç seri numaralı satırlar), M_CDP_STOK 18 MB (472 bin hücre). "Önceki satırla aynı alan = ~" kodlaması arşivleri
  ~4 kat küçültür (Trabzon 5,8 MB → 1,4 MB). Öneri: PAKET2 (geri uyumlu okuyucu) — kullanıcı onayı bekleniyor.

## Genel denetim (20.10.2026) — tüm birimler
- **Yükleme kartı**: frmSplash başlıksız kart, AKSORA penceresinin ortasında (GetWindowRect + SetWindowPos). Uzun döngülerde `AKSORA_BEKLE_NABIZ` (0,5 sn'de bir DoEvents, Interactive=False). Admin çizim aşamasına ara ilerlemeler eklendi. Kart Alt+F4 ile kapanmaz (QueryClose). Kart yüklenirken görsel hatası MsgBox çıkarmaz (`AKSORA_BEKLE_KIPI`).
- **Admin**: ADM_AC sırasında hesaplama elle (sonda geri). Önceki aramanın yeniden çalışması ekranı yarıda açmıyor (ScreenUpdating/Cursor ADM_AC'ye bırakılır). Tek başına arama kartla açılır. C4/C7 olay kilidi hata olsa da açılır. Giriş günlüğü 255 karakter sınırına göre kesilir (önceden 7-8 kayıttan sonra hiç yazılmıyordu). Demontaj "EKSİK GİRİŞ" aynı grup için bir kez sayılır (önbellek anahtarı DEM5). Kritik kullanımı olmayan il "%100 TAMAM" yerine "– / KULLANIM YOK". En iyi/en kötü rozeti yalnız gerçek fark varsa.
- **Defter**: `I_DEFTER_GUNCELLE` artık önbelleği silmiyor (imza zaten koruyor; il geçişleri hızlandı). Gün imzası en son yazılır.
- **İlçe**: SAP'den ilçe deposu satırı gelmezse ve önceki kayıtta o tesis varsa boş sonuç kaydedilmez (yoksa tüm stok "kullanıldı" sayılırdı). Sayfalar elle hesap/olaylar kapalıyken kurulur. İSG anahtarı Y_KOD_NORM ile de aranır. Aylık sayfada tek/eşit ilçede "EN KÖTÜ" gösterilmez.
- **Yatırım**: Hakediş tasarım kartı ve yazdırma alanı proje yüklenince tabloya göre yeniden konur (PDF kesik çıkmaz). Teminat "Sahada (gün)" kırmızısı yalnız sayı > 45. 0,75 gibi proje fiyatları okunur. Hata değeri olan teminat kaydı sayfayı kırmaz. Giriş kartında ara ilerlemeler; il seçilmemişse uyarı kart açılmadan.
- **Müşteri** (davranış/tasarım aynı): il yükleme yarıda kalırsa önceki il geri yüklenir (sayfalar iki ile bölünmez). Dağılım seçiminde olay/ekran ayarı hatada kapalı kalmaz. Tüm iller güncellemesi SAP kapalıysa baştan uyarır; sonda girişi yapılmış il geri yüklenir.
- **Malzeme**: CDP/EPP/sayfa güncellemesi SAP bağlantısı yoksa "tamamlandı" demez (C2 tarihi değişmediyse uyarı). Açılışta kapak tarihi hatası açılışı durdurmaz. İl geçişi arşiv kontrolü dizi ile (hücre hücre değil). Tüm stok güncelleme hata kodu kaybolmuyor.
- **Bekleyen öneriler (uygulanmadı)**: Müşteri aramasında Türkçe i/İ katlama; Malzeme `AKSORA_CANLI_STOK_IL` önbelleği; tek sistem güncellemesinde 4 kez CalculateFull; Teminat kurulumu toplu formül; ay geçişinde boşluk günleri yeni aya yazılıyor (her ayın son günü güncelleme önerilir).

## Tüm birimleri güncelle kartı + panel durumu (20.10.2026)
- Admin TÜM BİRİMLERİ GÜNCELLE artık birim panellerine gitmez: kart açılır, Excel penceresi `LockWindowUpdate` ile çizilmez (ekranda yönetim merkezi kalır), sonda kilit açılır, birim panelleri gizlenir, ADM_AC.
- Kart: birim (1/4), il sayısı (durum metinlerinde geçen il adları) ile dolar; son işlem ve geçen süre. Durum kaynakları: CORE_DURUM, AK11_DURUM, AKSORA_MUSTERI_TOPLU_DURUM, Y_DURUM, I_DURUM, NABIZ → `AKSORA_ADMIN_DURUM`.
- Yatırım/İlçe ana panel kartları: aynı gün GÜNCEL, gün değişince "1 gündür / N gündür güncellenmedi" (`Y_KART_TARIH`). Malzeme/Müşteri zaten tarih bazlıydı.
- İlçe STOK/DEMONTAJ: B6/I6 tarih kutuları kilitsiz (koruma uyarısı veriyordu); sayfa bir sonraki kurulumda düzelir.

## Tüm birimler: ekran sabit + DURDUR (20.10.2026)
- LockWindowUpdate yetmedi (Müşteri M_PencereKilit 0 ile kilidi açıyor). Çözüm: `AKSORA_ADMIN_TOPLU` iken birimler panel seçmez: Y_SAYFA_SEC, Y_PANEL_GOSTER, AKSORA_ANA_PANEL_GOSTER (yalnız yeniler), AKSORA_MUSTERI_PANEL_ON/GOSTER (yalnız yeniler), Malzeme toplu başı/sonu ANA PANEL seçimi.
- Kart başlığı baştan "AKSORA" (sonradan Caption değiştirmek başlık çubuğunu geri getiriyordu).
- DURDUR: kartta lblImza düğme olur (`AKSORA_BEKLE_AC(..., durdur:=True)`), tıklanınca `AKSORA_ADMIN_IPTAL`. Birim il döngüleri her il başında bakar; çalışan il biter, kalanlar ve kalan birimler atlanır.

## Hız/boyut: ortak SAP verisi + PAKET2 (20.10.2026)
- **Ortak MB52**: Yatırım ve İlçe, Malzeme ile birebir aynı MB52 PANEL raporunu il il çekiyordu (`AKSORA_YATIRIM_MB52` = `AKHZ_SISTEM_OKU` tek tesis). Malzeme "5 il tek sorgu" verisi `mOrtakVeri` olarak saklanır; `AKSORA_ORTAK_AKTIF` iken `AKSORA_YATIRIM_MB52` bu veriyi döndürür (Y_DONUSTUR tesise göre süzer). 30 dk tazelik, yeni okuma başında sıfırlanır.
- Malzeme "Bütün illeri güncelle" hatasız biterse `AK_ORTAK_ZINCIR` Yatırım+İlçe'yi aynı veriyle tamamlar (kart, panel değiştirmez, özet Malzeme mesajına eklenir). Admin sırası: Malzeme → Yatırım → İlçe (ortak) → Müşteri.
- **PAKET2**: alan önceki satırın aynı alanıyla aynıysa "=". Okuyucu PAKET1/PAKET2; yazıcı PAKET2 yazar, geri okuyup her alanı karşılaştırır, tutmazsa PAKET1. Açılış bakımı `_MA_/_YA_/_IA_` PAKET1'leri çevirir. Base'de 40 arşiv çevrildi (24 → 6,6 MB metin), `fx/verify_p2.py` ile birebir doğrulandı; kullanılmayan paylaşımlı metinler atıldı (optimize_base). Ham boyut 97 → 75 MB, dosya 19,1 → 15,8 MB.

## DERS: PAKET2 satırı "=" ile başlamamalı (20.10.2026)
- Müşteri arşivi yazılırken takıldı: blok hücresi "=" ile başlayınca Excel değeri formül olarak almaya çalışıyor. Kural: ilk alan (sütun 1) hiç "=" yapılmaz; yazıcı "=" ile başlayan bloğu reddeder. optimize_base aynı kural. Paket okuma döngüsüne NABIZ eklendi. Müşteri adımları teşhis dosyasına saatli yazılıyor.

## DERS: büyük arşivde doğrulama veriyi kopyalamamalı (20.10.2026)
- Müşteri CDP arşivi (17–37 bin satır) yazılırken Excel donuyordu; ilk doğrulama (PK_DOGRU) veriyi tam dizi olarak geri okuyup birkaç kez kopyalıyordu (ByVal + OKU sonucu). Yeni PK_YAZ: ByRef, yazılan hücreleri geri okuyup satır satır metin karşılaştırır. Canlı çalıştırmada her arşiv <1 sn, doğrulama hatası 0. Akış: Malzeme+Yatırım+İlçe ~35 sn, Müşteri il başına 28–47 sn.

## DERS: XML satır sonu normalleştirmesi (20.10.2026)
- Excel kayıtlı metinlerde CRLF var; XML okuyucu CRLF ve tek CR'yi LF okur. optimize_base bunu yapmadığı için PAKET2 satırlarının sonunda CR kaldı; blok sonundaki tek CR Excel'de fazladan satır oldu → Rize arşivi "satır sayısı tutmuyor" ile yüklenemedi (Giresun'dan sonra takılma). Düzeltme: `xmlnorm` (optimize_base + verify_p2), VBA okuyucu boş satırı atlar, müşteri il geçişi hataları teşhise yazılır.

## Doğrulandı (20.10.2026 13:53)
- Tüm birimler 4 dk 10 sn (önce 7 dk 10 sn): Malzeme 44 sn (CDP 18,1 + EPP 10,8 tek sorgu), Yatırım il başına 0,3–1,1 sn, İlçe 0,6–1,1 sn (ortak veri), Müşteri il başına 26–51 sn (toplam ~3 dk 6 sn; en büyük kalem).

## ATIL STOK (Malzeme birimi, 20.10.2026)
- Kural: 24 ay sarf (MB51 hareket 201/221/241/261/281/291) görmeyen malzeme ATIL; 18–21 ay "6 AY KALDI", 21–24 ay "3 AY KALDI". Gün sayacı: atıla kalan gün / atıl gün sayısı (bugüne göre, SAP'siz).
- Kapsam: illerin CDP+EPP **D001–D999 ve K001–K999** depoları, üretim yeri toplamı. HURDA ve ÇIKMA/BDSZ partiler, yüklenici (Z) depoları hariç. Stok kaynağı Malzeme birimi (canlı il: CDP_STOK/EPP_STOK/İLÇE_STOK; diğerleri `_AKLIVE_<il>_CDP/_EPP/_ILC`).
- SAP okuma (`AKSORA_ATIL_SAP_OKU`, sayfadaki "SAP'DEN OKU (MB51)"): sistem başına 5 üretim yeri; malzeme listesi panodan (Win32 pano, "Panodan yükle" btn[24], ilk değer kontrolü), depo D/K aralığı, düzen `/ATIL`. Adımlar: 1) son 6 ay, sistemin tüm aday malzemeleri (diğer illerin talebi de buradan) 2) bulunamayan adaylar 24–6 ay 3) kalanlar 60–24 ay (yoksa "5 yıldan eski"). "Kayıt yok" yalnız SAP mesajı açıkça söylüyorsa boş sayılır; diğer her hata okumayı durdurur, eski kayıt korunur. Açık olmayan sistemin eski satırları korunur. DURDUR çalışır (`AKSORA_ATIL_CALISIYOR`).
- Kayıt `_ATIL` (gizli): D1 CDP / E1 EPP okuma tarihi; satır: sistem, ÜY, malzeme, son sarf, son 6 ay sarf, kaynak (SAP/STOK), muhtemel sarf tarihi, SAP stoğu, son stok, son stok tarihi.
- Stok karşılaştırma: yeni stok verisi gelen ilde ÜY toplamı önceki miktardan düşükse "DÜŞTÜ (muhtemel sarf)" → sayaç o stok tarihinden yeniden başlar (≈ işaretli). Malzeme "Bütün illeri güncelle" sonunda ve sayfa her açıldığında çalışır.
- Sayfa `ATIL_STOK` (M_CDP_RAPOR başlığından kurulur, BX1 `AKSORA_ATIL_V1`): 4 KPI (ATIL / 3 AY / 6 AY / transfer önerisi), il düğmeleri (TÜM İLLER + 5 il), kategori özeti (KABLO, DİREK, PANO, HÜCRE, TRAFO, DİĞER), İL KARNESİ (atıl oranı = atıl değer / kapsam stok değeri) ya da tek ilde DEPO KARNESİ, TRANSFER ÖNERİLERİ (aynı sistem, son 6 ayda en çok sarf eden il; en fazla 100), tüm riskli malzemeler (filtreli).
- ANA PANEL'de "ATIL STOK" düğmesi (`goAtilStok`, goStokKarsilastir kopyası; 8 düğme B25:T26).
- MB51 sonucu klasik liste gelir (ızgara yok): dışa aktarma menüden (Liste > Kaydet > Dosya > Sekmeli metin). Sütunlar başlıktan (kısaltmalar dahil), olmazsa içerikten bulunur (tarih: değerlerin %90'ı tarih; miktar: sayısal ve eksi işaretli). Başlık + 3 örnek satır ve seçilen sütunlar teşhise yazılır.
- Sistemler ayrı okunur: biri hata verirse tamamlanan sistem kaydedilir, diğerinin eski kaydı korunur. `/ATIL` düzeni her SAP sisteminde (CDP ve EPP) ayrı kaydedilmeli; yoksa açık mesaj verilir. İlk CDP okuması (20.10.2026): 7154 son 6 ay 27.130 satır / 13 sn, 2-3. adımlar <1 sn.
- İlk tam okuma (CDP+EPP, 20.10.2026): 8 dk 26 sn. ATIL 2.435 kalem / 637,6 mn TL; 3 AY 103 kalem / 2,73 mn TL; 6 AY 63 kalem / 3,53 mn TL; transfer önerisi 691 kalem.
- Sayfa 3 sekmeli (BX21): ÖZET (kategori + karne), TRANSFER ÖNERİLERİ, ATIL LİSTESİ; tek bölüm gösterilir, sekmede filtre. Transfer: alacak il = son 6 ayda kullanan iller içinde eksiği (6 aylık kullanım − stoğu) en büyük olan; YÜKSEK = stoğu 6 aylık kullanımından az, gönderilecek = min(eldeki, eksik); DÜŞÜK = stoğu yetiyor, gönderilecek = min(eldeki, 6 aylık kullanım). Sütunlar: Nereden, Sistem, Malzeme, Tanım, Depo, Eldeki miktar, Değer, Durum, Nereye, Gönderilecek, Öncelik, Neden (tek cümle).
- Admin bölümü sonra eklenecek.

## Kullanıcı dosyası 20.10.2026 (317c039b, yeni baz `up23`)
- Kod bizim son sürümle aynı (yalnız VBE büyük/küçük harf), yeni Sayfa95 = ATIL_STOK (kodsuz), frmSplash VB_Base değişti → sonraki derlemede raw'a up23'ten alınır.
- 19,0 MB: Müşteri ekran ili Trabzon (M_CDP_STOK 38.246 satır), Excel görselleri sıkıştırmadan saklıyor, VBA derleme önbelleği. Sonraki derlemede optimize_base ile küçülür.
- Müşteri il başına 25–55 sn; en büyük kalem IQ09 (Trabzon 14 sn sorgu + 19 sn liste).
- ATIL ilk sonuç şüpheli: il atıl oranı %72–81, stoktaki kalemlerin %45'i "5 yıldır sarf yok" (2.252). Muhtemel neden: sarf yalnız D/K depolarından sayılıyor; yatırım ambarından yükleniciye (Z) giden ve orada sarf edilen malzeme hiç sarf görmemiş sayılıyor (ör. Artvin EPP D002 Yatırım Ambarı %100 atıl). Ayrıca yeni girip henüz sarf edilmemiş malzeme "5+ yıl" görünür; iptal hareketleri (202/262…) düşülmüyor. Karar kullanıcıda (doğruluk testi).

## 21.10.2026
- Kullanıcı Artvin'deki tüm "5 yıldır sarf yok" kalemlerini SAP'de kontrol etti: gerçekten sarf yok (doğru).
- ATIL_STOK, KAPAK'tan Malzeme'ye girişte (AKSORA_GIRIS_AC) etiketi doğruysa görünür yapılır (önceden yalnız düğmeyle açılıyordu).
- Baz up23'e taşındı (optimize_base): 19,0 → 16,0 MB; _ATIL, ATIL_STOK, M_CDP_STOK verisi birebir, 40 arşiv fark 0.

- Kullanıcı yeni dosya verdi (baz `up24`, 19,7 MB); ATIL sekme düzeltmesi bunun üzerine kuruldu → 16,2 MB, veri birebir. ATIL için yetkin kişiyle test yapılacak, sonuç kullanıcıdan gelecek.

## Arka plan modu (21.10.2026)
- Kullanıcı: "Tüm birimleri güncelle" sırasında Excel kilitli, küçültülemiyor (sebep: tüm iş boyunca Interactive=False + SAP sorgusu sırasında mesaj işlenmiyor).
- DURDUR'lu kartta (tüm birimler, ATIL SAP okuması) yanına **ARKA PLANA AL** (lblAltBaslik): `AKSORA_ARKA_PLANA_AL` Excel'i simge durumuna alır, Interactive=True (görev çubuğundan geri açılabilsin). NABIZ/ILERLE arka planda giriş kapatmaz; kullanıcı pencereyi açarsa `ARKA_PLAN_KONTROL` arka planı bitirir, tüm birimlerde girişi yeniden kapatır. İş bitince `AKSORA_ARKA_PLAN_BITIR` pencereyi eski haline getirir, sonra sonuç mesajı.
- Kart düzeni (durdurlu): lblDurum iki satır (adım + "Birim n/4 • Geçen süre"), çubuk 92, düğmeler 114 (DURDUR kırmızı, ARKA PLANA AL mavi), yükseklik 182.
- Küçük pencerede pencere ayarı yapan yerlere `AKSORA_PENCERE_KUCUK` koruması: V13_GORUNUM, RAPOR/STOK_GORUNUM_TEK, MUSTERI_DUZEN, ANA_PANEL_UYGULA.
- Müşteri IQ09'u 5 il tek sorguya indirmek ölçüldü: sabit maliyet il başına ~5 sn, kazanç ~20-30 sn; risk/kazanç düşük → yapılmadı.

## 21.10.2026 eklemeler
- Arama (Müşteri, Malzeme il araması, Admin): `AKSORA_ARAMA_DESEN` (modAKSORA_SISTEM) ile `Like` deseni; i/I/İ/ı, s/ş, g/ğ, u/ü, o/ö, c/ç ve büyük/küçük harf aynı ("iletken" → ILETKEN, "şalter" → SALTER).
- Admin: il karşılaştırmasının altında **ATIL STOK** tablosu (`ADM_ATIL_CIZ`, ATIL_HESAPLA ile aynı hesap; _ATIL yoksa çizilmez): il başına ATIL/3 AY/6 AY kalem-değer, atıl oranı, transfer önerisi, yüksek öncelik, muhtemel sarf, değerlendirme.
- ATIL transfer sekmesinde **PDF ÇIKTI** (`AKSORA_ATIL_TRANSFER_PDF`): yatay, 1 sayfa genişlik, başlık satırı tekrar, filtre uygulanmışsa yalnız görünen satırlar; AKSORA kayıt klasörüne `ATIL_TRANSFER_<il>_<tarih>.pdf`, açılır.
- Teminat kurulumu: satır formülleri sütun bazında tek seferde (FormulaR1C1), sonuç aynı.
- ANA PANEL'e ATIL kartı eklenmedi (kilitli panel tasarımı; ATIL STOK düğmesi var).

## DERS: Excel simge durumundayken güncelleme yapılmaz (21.10.2026)
- İlk ARKA PLANA AL Excel'i küçültüyordu: ekrandaki ilin (Artvin) stok sayfalarında depo açılır listesi (DropDown ListFillRange) doldurulamadı → Hata 1004, her birimde Artvin güncellenemedi, Malzeme'de "Görünüm yenilenemedi". Kullanıcı dosyayı kaydetmedi.
- Yeni ARKA PLANA AL: pencere küçültülmez, kart + Excel z-sırasında en alta gönderilir (SetWindowPos HWND_BOTTOM, NOACTIVATE); Interactive kapalı kalır. Bitişte görev çubuğu yanıp söner (FlashWindowEx), pencere zorla öne getirilmez.
- Güvence: iş sürerken Excel küçültülmüşse NABIZ/ILERLE'de geri açılıp arkaya gönderilir (`ARKA_PLAN_KORU`); açılır liste dolduran yerler önce `AKSORA_PENCERE_GERI_AC` çağırır (tek birim güncellemelerinde de).

- Doğrulandı (21.10.2026 11:49): yeni ARKA PLANA AL ile tüm birimler hatasız, 4 dk 21 sn (Malzeme 42 sn, Yatırım/İlçe il başına 0,3–1,4 sn, Müşteri il başına 30–50 sn).

## 21.10.2026 (baz `up25`)
- Kullanıcı: arka plan modu sorunsuz (alta al / aç / tekrar alta al). Admin ATIL tablosu görüldü.
- Müşteri aramasında "iletken" sonuç vermedi: Müşteri verisinde (5 il, arşiv dahil) ILETKEN yok → doğru. Test kelimesi: "trifaze" (2.616 satır).
- ATIL sayfasından AKSORA filigranı (AKSORA_CDP_V1_FILIGRAN) silinir.
- Transfer PDF yeniden: geçici çıktı sayfası; özet (öneri sayısı, toplam değer, YÜKSEK/DÜŞÜK, nereye il dağılımı), kurum yazıcısına uygun açık gri başlık + siyah yazı, dolgusuz satırlar, Sistem sütunu yok; yalnız filtrede görünen satırlar; sayfa sonra silinir.
- "Neden" kısaldı: "Son 6 ayda X kullandı; deposunda Y var / hiç yok." (il adı Nereye sütununda).

## KURAL: AKSORA'da donmak yok (21.10.2026, kullanıcı)
- Bekleten her işte yükleme kartı (AKSORA_BEKLE_AC) açılır, döngülerde NABIZ/ILERLE. Eklendi: Müşteri/Yatırım/İlçe malzeme araması (il il ilerleme), Malzeme il araması, ATIL transfer PDF, Hakediş PDF, Müşteri stok sayfasına geçiş. MsgBox'tan önce kart kapatılır.
- Transfer PDF miktarı AKSORA_V13_MIKTAR_BICIMI (kesirsizde virgül yok).
- Kullanıcı onayladı: transfer PDF özet + yazıcı renkleri uygun.

- Müşteri/Yatırım/İlçe araması yalnız seçili ili tarar (BX21; TÜMÜ = 5 il). Kapsam BX28; arama sonrası başka il seçilirse arama o kapsamla yenilenir, TÜMÜ ile yapılmışsa yalnız filtre. Kullanıcı: PDF miktar biçimi ve arama kartı onaylı.

## ATIL kuralı 2 (giriş yaşı + yüklenici) — 2026-10-21
- Kullanıcı kuralı: atıl üretim yeri bazında; yüklenici depoları ve yüklenici sarfı dahil. Örnek: 7155 D001 50 + 7055 D002 10, yüklenici 7055'ten 5 alıp sarf etse atıl toplamı 55 kalır.
- Hesap: eldeki stok FIFO ile girişlere dağıtılır (en yeni giriş önce kalır). 24 aydan eski girişten kalan (ya da girişi bulunmayan) kısım ATIL; 21–24 ay 3 AY KALDI; 18–21 ay 6 AY KALDI. Kalem birden çok basamakta olabilir (x(19) değer, x(20) miktar).
- Giriş türleri: 101/102, 105/106, 501/502, 561/562 işaretiyle; 301/303/305/315 yalnız artı, 302/304/306/316 yalnız eksi (alan taraf). 311 (aynı ÜY) yaşı sıfırlamaz. Sarf: 201/221/241/261/281/291 ve iptalleri.
- Kapsam: D###, K###, M###, BO## ve yüklenici stok sayfasındaki depolar (YÜKLENİCİ_STOK / _AKLIVE_<il>_YUK; mAtilYuk). HURDA ve ÇIKMA/BDSZ hariç.
- SAP: ÜY başına 3 MB51 sorgusu (/ATIL düzeni; İTü ve DpYr sütunları gerekli; depo filtresi yok, MATNR ve BWART panodan): 1) 24 ay girişler, 2) stoğu girişleri aşanlar için 24–60 ay son giriş, 3) 6 ay sarf (birleşik malzeme listesi).
- _ATIL etiketi ATIL2: sis | ÜY | mal | girişler "yyyymmdd=miktar;…" | eski giriş | 6 ay sarf | son sarf | SAP stoğu | son stok | son stok tarihi. ATIL1 kaydı okunmaz; sayfa "kural değişti, yeniden okuyun" der.
- Sonraki stok verilerinde artış stok tarihiyle yeni giriş katmanı; okumadan sonra stoğa giren kalem tamamı yeni giriş.
- Liste/transfer/PDF: Riskli miktar ve Riskli değer (18 aydan eski kısım). "Muhtemel sarf" yerine "Stoğu düşen".

## SAP malzeme hafızası (MAT) düzeltmesi — 2026-10-21
- Sorun: ATIL MB51'de MATNR'ye yapıştırılan listenin ilk değeri SAP hafızasına (MAT) yazıldı; sonraki MB52 stok okumaları tek malzemeyle çalıştı (Artvin CDP 3, EPP 1 satır), ilçe adımında PANEL düzeni seçilemedi.
- Çözüm: AKSORA_SAP_MALZEME_TEMIZLE (TOPLU_GUNCELLEME, Public) her StartTransaction "MB52" sonrasında MATNR-LOW/HIGH temizler (9 yer). ATIL bitişinde MB51 açılıp MATNR boşaltılır, Enter.
- EPP MB51 "Herhangi bir düzenek bulunamadı": /ATIL yoksa varsayılan düzenle bir kez yeniden denenir (teşhise yazılır).
- ATIL transfer tablosu ve PDF'de Depo sütunu tüm depoları miktarıyla gösterir ("D001 (3), K002 (5)"); uzun listede satır yüksekliği açılır.

## İlçe PB depoları doğrulaması — 2026-10-21
- Artvin CDP ilçe verisinde PB01 (planlı bakım) çıktı; ILCE_V2_UY_DOGRULA yalnız K### kabul ettiği için zincir durdu. Doğrulama artık AKSORA_IL_DEPO_GECERLI(depo, "ILCE") kullanır (K001-K999 + PB01-PB999). Admin ADM_DEPO_TUR PB'yi ilçe sayar.

## EPP MB51 liste gösterimi — 2026-10-21
- EPP'de MB51 liste gösterimi hiyerarşiğe dönmüştü; /ATIL düzeni düz listede kayıtlı olduğundan SAP "Herhangi bir düzenek bulunamadı" dedi, varsayılan düzen hiyerarşik geldi (İTü/depo yok). ATIL_DUZ_LISTE sorgudan önce seçim ekranındaki "Düz liste" radyo düğmesini metninden bulup seçer (teşhise yazar).

## ATIL KANIT düğmesi — 2026-10-21
- MC50 bozuk; kurumda atıl MB51'den "2 yılda sarfı olmayan" çekilerek karşılaştırılıyor. AKSORA kuralı (giriş yaşı) farklı, o yüzden her kalem için kanıt: AKSORA_ATIL_KANIT (Public). TRANSFER ve LİSTE görünümünde "SEÇİLİ MALZEME: KANIT" düğmesi; seçili satırın il/sistem/malzemesinden _ATIL kaydı ve stok okunur, son 24 ay girişleri (yeni→eski, ilk 8), her girişin stokta kalan payı ve yaşı, karşılanmayan kısım, 2-5 yıl son giriş, 6 ay sarf, MB51 kontrol bilgisi MsgBox'ta (1024 sınırı).
- Baz up26'ya taşındı (kullanıcının ATIL2 okuması yapılmış dosyası, 52feed3a; optimize_base --img): 152 sayfa verisi birebir, 40 arşiv fark 0, _ATIL ATIL2.

## Genel denetim ve düzeltmeler — 2026-10-22 (baz up28)
- Telif: sahip adı kodda düz metin değil (KORUMA IMZA_KOD, AKSORA_SAHIBI() fonksiyon); kartlarda AKSORA_TELIF_ETIKET ile soluk "® Burhan SERDAR" (frmSplash'a çalışma anında lblTelif).
- ATIL: girişler ilin stok tarihine kadar okunur (stok tarihinden sonraki giriş çift sayılmıyordu); 313/315 giriş değil; ALV boş hücre yalnız malzeme hücresi boşken doldurulur; KANIT/PDF okuma sırasında çalışmaz; SAP okuma öncesi ve STOK_KONTROL'de kart; ATIL_GOSTER hatada sayfayı korur; Admin'de verisi olmayan il "Veri yok".
- Yatırım: Y_PROJE_KARSILASTIR IIf → açık dallanma (eksik anahtarda Type mismatch); Teminat kurulumu başında BX25 silinir; Teminat/Hakediş elle hesaplamada kurulur; Y_/I_ sayfası yoksa güncelleme hata vermez; tek il Yatırım/İlçe güncellemesinde kart.
- Stok: yedek MB52 okuyucularında (EPP, İLÇE EPP, Yüklenici x2) MATNR temizlenir; bütün iller sonunda raporlar canlı il ile yenilenir; BIRLESTIR sessiz modda hatayı iletir; stok sayfası SAP düğmesinde kart; il geçişinde C3/C4 SUBTOTAL formülü korunur.
- Müşteri: fiyat listesi tüm _MA_ il kayıtlarına da uygulanır; Excel aktarım hatasında DisplayAlerts geri; Excel işlemleri yalnız SAP Excel'e aktardıysa kapatılır; IQ09/MB52 liste yoksa SAP mesajında 5 sn'de çıkılır; hatalarda sayfa koruması; fiyat yüklemesinde kart; tek satırlı tabloda ÇIKTI (SpecialCells Intersect).
- Bilinçli dokunulmayan: MS_Birlestir "çift sayım" bulgusu (gerçek veride AKSORA = MB52, fark 0); ilçe yedek yollarında PB sorgusu (hızlı yol kullanılıyor); performans iyileştirmeleri (banding, KARS yazımı) ileride.

## İlk SAP erişimi ve bozuk modül kaynağı — 2026-10-22 (baz up29)
- Her oturumda bekleme sonrası ilk ekran erişimi düşüyor (ATIL ilk sorguda RFLAT_L "Object doesn't support this action", Tüm birimler ilk hızlı okumada 619; ikinci deneme sorunsuz). AKHZ_SISTEM_OKU 619'da 1 sn bekleyip bir kez yeniden dener, açık wnd[1]'i kapatır, düşülen adımı [adim] ile yazar. ATIL_DUZ_LISTE iki deneme.
- Kullanıcının 22.10 09:05 kaydında modAKSORA_TASARIM kaynak akışı bozuktu (23. sıkıştırma parçası başlığı bozuk; Excel p-code ile çalışıyordu). Derleme kaynağı raw'dan yeniden yazar. extract.py bozuk modülde durur: tek tek atlayarak karşılaştır.

## Ekran incelemesi düzeltmeleri — 2026-10-22 (baz up30)
- Filtreli listede End(xlUp) gizli satırları atlıyordu: depo filtresi açıkken il geçişi arşive eksik kopyalıyor, Admin ATIL Artvin 249 (doğrusu 259) sayıyordu. Tüm `X.Cells(X.Rows.Count, c).End(xlUp).Row` (111 yer) → AKSORA_SON_SATIR(ws, kolon) (Find xlFormulas, gizli satırlar dahil; MUSTERI_ISLEMLER).
- modAKSORA_TASARIM kaynağı kullanıcı kaydında iki kez bozuldu (gömülü 5,3 KB base64 imza PNG'si bölgesinde Excel sıkıştırması). PNG kitabın gizli adlarına taşındı: _AKS_IMZA_001..027 (200 karakterlik parçalar); T_CiktiImzaResmi adlardan okur. Derlemede tools/imza_names.py (scratchpad fx/) final'e adları ekler.
- Y_MALZEME_ARA alt başlığı "Yatırım (D002, D011) + Yüklenici Stokları", I_MALZEME_ARA "CDP + EPP İlçe Depoları"; ilçe notlarına PB eklendi; stok sayfalarında A sütunu en az 12 (Üretim Yeri başlığı kesilmesin).

## Birim şifreleri — 2026-10-22
- ŞU AN KAPALI: BIRIM_SIFRE_AKTIF = False (BIRIM modülü başı). Paylaşımda True yapılacak.
- Hub'daki "İşlemlere geç" (AKSORA_BIRIM_MALZEME/YATIRIM/MUSTERI/ILCE) önce BIRIM_SIFRE_KONTROL: birimin şifresi (SIFRE_MALZEME/YATIRIM/MUSTERI/ILCE sabitleri, BIRIM modülü başı) ya da yönetici şifresi. Doğru giriş oturumda hatırlanır (mBirimAcik); yönetici şifresi tüm birimleri ve Admin'i açar. Kutu ADM_SIFRE_SOR (yıldızlı, Şifreyi göster), başlık "AKSORA - BIRIM GIRISI". Açılışta yalnız HUB görünür, diğer birim sayfaları VeryHidden.

## Admin canlı il ve CEO kataloğu — 2026-10-22
- ADM_CANLI_IL birimin ekrandaki ilini A5 başlığından okuyordu; Yatırım/İlçe il değişince A5 güncellenmiyor → Artvin satırına Gümüşhane yatırımı (89,2 mn) ve Trabzon ilçe/İSG (2.595) yazılıyordu. Artık önce verinin ÜY kodundan (A7:A40, ATIL_TESIS_IL), sonra BX19/BX20/A5.
- Doğru değerler (22.10 10:38 dosyası, çevrimdışı): Artvin İSG zimmet 1.042 adet/253 kalem, 5 il toplam 4.297; Artvin yatırım 119,99 mn TL; Yatırım toplam ≈1.073 mn; İlçe 9.911 satır/160 depo/39,4 mn; Müşteri 145.108 satır; atıl il sayıları ATIL sayfasıyla aynı (Artvin 259).
- Katalog: katalog/AKSORA_Katalog.pdf (13 sayfa A4 yatay; HTML scratchpad/katalog → Chromium print-to-pdf). Yalnız doğrulanmış rakamlar; hatalı Admin ekranları kullanılmadı.

## 02.11.2026 – Demontaj yalnız 6 kritik malzeme + İlçe birleşik stok/rapor
- Taban: kullanıcının 02.11 dosyası (up31, kod aynı; yalnız veri yeni).
- Demontaj kapsamı: yalnız TRAFO, PANO, ARMATÜR, KABLO, DİREK, HÜCRE. `I_DEM_DISI` kritik olmayan her malzemeyi
  "Kritik değil (demontaj kontrolü dışı)" döndürür → günlük STOK/DEMONTAJ ekranı, Yönetim Merkezi günlük demontaj,
  aylık defter (yeni kayıtlar) bu kurala uyar. `I_AY_HESAP` eski defterdeki kritik dışı kalemleri hesaba katmaz.
- ARMATÜR yeni kategori: mal grubu `M030103*` (ARMATUR SODYUM/LED...); mal grubu yoksa tanım "ARMATUR " (BORUSU hariç).
  Sıra 6, renk sarı-kahve. Karne dizisi: 4..9 kategori eksikleri, 10 = uyum (eskiden 9).
- I_AYLIK_DEMONTAJ 12 sütun (kategori tablosunda "Fireli çeşit", fire listesinde "Kullanılan değer", karnede "Armatür eksik").
  Admin karnesinde uyum tek sütun, 6 kategori sütunu; il özetinde "Pano / hücre / trafo / armatür eksik". Admin önbellek imzası DEM6.
- İlçe birimi: `I_BİRLEŞİK_STOK` + `I_BİRLEŞİK_RAPOR` (CDP + EPP ilçe depoları tek tabloda, `I_BIRLESTIR`).
  I_ADLAR/I_TUR/Y_BASLIK eklendi; sayfalar birime girişte Y_KUR ile kurulur (güncelleme önce gelirse I_STOK_YAZ kurar).
  Panel hızlı erişim 6 düğme (IP_V5).

## 02.11.2026 – Stok sayfaları A-Z + EXCEL'E AKTAR
- Tüm stok sayfaları (Malzeme CDP/EPP/STOKLAR/İLÇE/YÜKLENİCİ, M_/Y_/I_ *_STOK) malzeme tanımına göre A-Z
  (sonra malzeme kodu, depo, seri no; Türkçe harf sırası). `AKSORA_STOK_DIZI_AZ` (TASARIM, Public) dizi sıralar;
  `Y_STOK_YAZ` yazmadan önce sıralar; `AKSORA_STOK_AZ ws` sayfa açılışında (AKSORA_CIKTI_DUGME) ve T_StokYap'ta
  sıralı değilse yerinde sıralar, filtre varsa `AutoFilter.ApplyFilter` ile yeniden uygular.
- "EXCEL'E AKTAR" düğmesi (btnExcelAktar, yeşil) ÇIKTI AL'ın solunda: `AKSORA_EXCEL_AKTAR` → T_CiktiSayim (mExcelAktar)
  → `T_ExcelYaz`: yeni görünür kitap, 1. satır başlık (sayfa | il | depo | stok tarihi), 2. satır hazırlama,
  4. satır sütunlar; miktar sayı + Birim ayrı, SAYILAN boş, FARK formülü; A-Z, çerçeveli, filtre, başlık dondurulmuş,
  A4 yatay sayfa düzeni (imza bandı yok). HURDA ve ÇIKMA/BEDELSİZ partiler çıktıdaki gibi alınmaz.
- EXCEL'E AKTAR düzeltme: miktar/SAYILAN/FARK Genel biçim (binlik ayraç yok, küsurat virgülle), üretim yeri ve
  malzeme kodu sayı (yeşil uyarı yok); yeni kitap açılınca UYGULAMA_MODU_KAPAT + şerit/formül çubuğu açılır (olaylar kapalı
  olduğundan Workbook_Deactivate çalışmıyordu). AKSORA'ya dönünce Workbook_Activate uygulama modunu açar.
- Çıktı al + Excel'e aktar: HURDA hiç alınmaz; ÇIKMA/BDSZ partide yalnız TRAFO (M0340), HÜCRE (M0314), PANO (M0327),
  DİREK (M0308) alınır (`T_CiktiDahil` / `T_CikmaKritik`; mal grubu yoksa tanımın başı).
- Excel'e aktar görünüm düzeltmesi: T_ExcelYaz artık AKSORA penceresine dokunmaz (UYGULAMA_MODU_KAPAT çağrılmaz);
  `AKSORA_EXCEL_GORUNUM` (UYGULAMA) yalnız şerit/formül çubuğu/sistem ayraçlarını açar. ThisWorkbook'ta
  `Private WithEvents mUygApp As Application` + `AKSORA_UYG_IZLE` (CallByName ile kurulur): AKSORA penceresi etkinleşince
  UYGULAMA_MODU_AC, başka kitapta AKSORA_EXCEL_GORUNUM. UYGULAMA_MODU_AC'de On Error en başa alındı.
- DONMA düzeltmesi (müşteri seri no): A-Z sonrası depo filtresi binlerce görünür parça üretir; SpecialCells(xlCellTypeVisible)
  + Areas Excel'i kilitliyordu. `AKSORA_GORUNUR(ws, ilk, son, sonKol)` (SISTEM): tek `SUBTOTAL(103,OFFSET(satır,ROW()-ilk,0))`
  dizi hesabı, olmazsa Rows().Hidden. Kullanan: T_CiktiSayim, T_CiktiListe, AKSORA_MUSTERI_BANT (zebra).
  YENİ KOD YAZARKEN stok sayfalarında SpecialCells(xlCellTypeVisible).Areas KULLANMA.
  AKSORA_STOK_AZ oturum önbelleği (mAzImza: satır sayısı, C2, A5, ilk/orta/son satır).
- Müşteri sayfa açılışı donması: AKSORA_STOK_AZ sayfa açılışında 5000+ satırlı sayfaları sıralamaz (18k satır 3-5 sn
  kartsız bekletiyordu). Müşteri sayfaları yazılırken sıralanır: MS_Yaz ve il kaydından yükleme (AKSORA_STOK_DIZI_AZ,
  4096 satırda bir AKSORA_BEKLE_NABIZ). Anahtarda Replace yalnız Türkçe harf varsa.
- Not: fx/audit_vba.py `anah()` gibi dizi parametreli imzalarda parametreleri kaçırır (yanlış "unknown").

## 02.11.2026 – Tam analiz (4 paralel inceleme) + Excel'e aktar masaüstü / toplu / demontaj
- Taban: up32 (kullanıcının 02.11 07:15 dosyası; kod aynı).
- YENİ: Excel'e aktar dosyayı masaüstüne "<depo> - <birim sayfa> - <stok tarihi>.xlsx" olarak kaydeder (AKSORA_MASAUSTU:
  WScript.Shell Desktop, OneDrive dahil; T_Kaydet ayni ad açıksa "(2)"). Başlık "I_" yerine "İLÇE YÖNETİM | ...".
- YENİ: "TÜM DEPOLARI AKTAR" (btnTopluAktar; yalnız I_*_STOK ve M_*_STOK): her depo ayrı dosya (depo adıyla) masaüstünde
  "AKSORA - <İl> - <sayfa> - <tarih>" klasörüne; yükleme kartı, sonunda klasör açılır (AKSORA_EXCEL_TOPLU_AKTAR,
  mDepoFiltre ile T_CiktiSayim).
- YENİ: İlçe STOK/DEMONTAJ ekranında "DEMONTAJ EKSİKLERİ EXCEL'E AKTAR" (ÇIKTI AL'ın sağı): B6/I6 tarihleri için
  I_DEM_EKSIK (ekranla aynı kural, yalnız 6 kritik) → her ilçe ayrı dosya + "00 - Tüm İlçeler" (AKSORA_EXCEL_TABLO, genel).
  Artvin 23.10→02.11 doğrulama: 9 ilçe, 22 kalem, 63.562 TL.
- DÜZELTMELER (analiz): 43 prosedürde "Resume Cikis + Cikis'te hata → sonsuz döngü" (Cikis başına On Error Resume Next,
  Err.Raise öncesi On Error GoTo 0; fx/resume_fix*.py). AKSORA_BEKLE_NABIZ kart yokken de DoEvents. Yatırım: SAP boş
  gelirse kayıt silinmez (2103). Teminat detay sıralaması Y_SIRALA (A-Z sonrası O(n²) donma). Admin arama çift sayım,
  eksi stok üçlü sayım (il için en yüksek birim). ADM_DEMONTAJ önce I_MAL_NOT + önbellek imzasına harita sayısı.
  I_KATEGORI tanım Y_NORM (DİREK/HÜCRE/ARMATÜR). Tüm birimler: EnableCancelKey kapalı. Bütün iller (ilçe/yatırım) kartla.
  Satır döngülerine NABIZ. ADM_AC hata sonrası koruma. Panel kartı ve canlı arşiv filtreden bağımsız sayı/değer.
  STOK_KARS KPI ve TEMA_GORUNUR_DEPO_SAYISI AKSORA_GORUNUR. İlçe yedek MB52 okuması K001..PB999 (PB depoları) + ilçe dışı atla.
  Workbook_Open: yardımcı adımlar ayrı korumalı, hata olursa AKSORA_BIRIM_ACILIS; SheetActivate birim hatası yakalanır;
  gizli KAPAK yerine birim seçimi; mUygApp dönüşte seçim ekranı görünümü. AKG_KOPYALA sonrası ApplyFilter.
  Müşteri bant (BW) eski fazla satırları silinir (dosya şişmesi). Excel'e aktar kod sütunları metin (baştaki sıfır korunur).
- Kalan maddeler de düzeltildi (aynı gün, ikinci tur):
  * Stok güncellemesi hata verirse (CDP/EPP/İLÇE/YÜKLENİCİ/STOKLAR) C2/C3/C4 eski haline döner, rapor eski veriden yeniden kurulur (eskiMeta).
  * ATIL_STOK_OKU: ilin kaynak sayfaları arasında 1 saatten fazla fark varsa tarih = en eski kaynak (kısmi yenilemede karşılaştırma yapılmaz).
  * AKG_KOPYALA: arşiv canlı sayfadan yeniyse canlı→arşiv kopyası atlanır (yeni SAP verisi korunur).
  * AKHZ_SAYI: tek nokta + 3 hane değilse ondalık nokta (1234.56); çok noktalı binlik.
  * ADM_VERI: başka ilin üretim yerine ait satır sayılmaz (ADM_IL_UYGUN, önbellekli).
  * I_KARNE: aynı kategoride farklı birim → -1 ("farklı birim"), toplanmaz.
  * BeforeClose: kaydetme sorusu AKSORA sorar; İptal'de uygulama modu kapanmaz.
  * Hata işleyicileri: Y_KORU + önceki EnableEvents (4 yer + I_AYLIK); KAPAK_KORU / KILITLE / GIRIS_AC önceki durum,
    açılışta MsgBox yerine Err.Raise / sessiz; yarım il seçim kopyası silinir (yeniKopya).
  * BİLİNÇLİ OLARAK AÇIK: boş kategori (0 satır) toplu güncellemede reddedilir; boş SAP okumasının veriyi silmemesi için
    korumadır. Gerçek boş kategori şu an hiçbir ilde yok.
- ed.py düzeltildi: önce kodla sonra yaz (cp1254 dışı karakter dosyayı 0 bayta indiriyordu).

## 02.11.2026 – İSG Zimmet Kontrolü Excel'e aktar
- Taban: up33 (kullanıcının 02.11 dosyası; yeni sayfa modülleri Sayfa96/97 = Y_KUR'un kurduğu I_BİRLEŞİK_* sayfaları).
- DİKKAT: build.sh artık `set -e -o pipefail`; yeni dosyada raw'da olmayan modül varsa build_bin hata verip durur
  (eskiden sessizce eski new_vbaProject.bin ile paketliyordu). Yeniden temel almada u*/src'de olup raw'da olmayanlar kopyalanmalı.
- I_İSG_MALZEMELERİ'nde "EXCEL'E AKTAR" (btnIsgAktar, ÇIKTI AL'ın solu) → AKSORA_ILCE_ISG_AKTAR: 1. sayfa "İlçe Durumu"
  (ekrandaki durum renkleri + TOPLAM), 2. sayfa "Malzeme Listesi" (ZİMMET YAPIN kırmızı, MANUEL SARF YAPIN mavi; ekrandaki
  filtre / ilçe seçimine uyar; miktar ve birim ayrı). İlçe seçiliyse yalnız o ilçe. Masaüstü: "<il|ilçe> - İSG Zimmet Kontrolü - <tarih>".
- TASARIM: AKSORA_EXCEL_SAYFA (sayfa düzeyi tablo) + AKSORA_EXCEL_BITIR (çok sayfalı kitap kaydet); AKSORA_EXCEL_TABLO bunları kullanır.

## 02.11.2026 – ALTIN KURAL: tutarlar "#.##0,00 TL"
- Kullanıcı kuralı (tüm sistemler, tüm Excel çıktıları): binlik nokta, kuruş virgül, sonda TL. VBA NumberFormat: "#,##0.00 ""TL""".
- AKSORA_EXCEL_SAYFA → T_ExcelBicim: "0" kod sütunu metin; "0.00" içeren TL'siz biçime otomatik " ""TL""" eklenir.
- İSG (2 sayfa) ve demontaj eksik değer sütunları TL; STOK_HAFIZA delta değer/özet TL (delta 12-13 miktar "#,##0.###").
- Taban: up34 (kullanıcının 02.11 dosyası, kod aynı).
- YENİ KOD YAZARKEN: her tutar hücresi/sütunu ve mesaj metni TL ile biter.

## 03.11.2026 – Müşteri yazımı 36 sn → 431 sn gerilemesi düzeltildi
- Teşhis: Müşteri ARTVIN "Sayfaya yazılıyor: M_CDP_STOK" 08:02:08 → 08:08:44 (önceden 1-2 sn). Sebep: AKSORA_GORUNUR'daki
  ws.Evaluate("SUBTOTAL(103,OFFSET(satır,ROW(...)-ilk,0))") 18 bin satırda dakikalar sürüyordu (M_STOK_KUR → AKSORA_MUSTERI_BANT).
  Bu sırada DoEvents olmadığından kartın ARKA PLANA AL / DURDUR düğmeleri tıklanamıyordu.
- Çözüm: AKSORA_GORUNUR: FilterMode yoksa hepsi görünür (anında); filtre varsa Rows().Hidden döngüsü + 2048 satırda NABIZ.
  OFFSET/SUBTOTAL dizi Evaluate'i büyük sayfada KULLANMA.
- ASIL SEBEP (2. teşhis, 327 sn sürdü): T_AzAnahtar(ByVal a As Variant ...) her satırda 18k×14 diziyi kopyalıyordu
  (≈4,5 milyar kopya). ByRef yapıldı. KURAL: büyük diziyi döngü içinde çağrılan fonksiyona ASLA ByVal geçirme.

## 03.11.2026 – Yükleme kartı güven / ARKA PLANA AL
- Tüm birimler 4 dk 18 sn'ye döndü (Müşteri il başına 30-50 sn).
- Kart takılması: SAP rapor çağrıları (IQ09 / MB52 / toplu okuma) tek çağrıda 10-19 sn Excel'i bloke eder; bu sürede
  VBA kodu çalışamaz, kart güncellenemez, tıklama sonradan işlenir. Kartta bu adımlarda "(SAP yanıtı bekleniyor)" yazar.
- ARKA PLANA AL: tıklanınca düğme yeşil "ARKA PLANDA • ÖNE GETİR" olur + bilgi metni; ikinci tıklama AKSORA_ONE_GETIR
  (HWND_TOP). Eskiden pencere z-sırasında alta gidiyordu, başka pencere yoksa hiçbir şey olmamış gibi görünüyordu.
- NABIZ eşiği 0,5 → 0,25 sn. Taban: up35.

## 03.11.2026 – SATINALMA KONTROLÜ (ME5A + ME2N) – ilk sürüm
- Kullanıcı tercihi: yanıtlarda açıklama yok, yalnız "Yap:" adımları.
- Yönetim Merkezi'nde atıl bölümünden sonra "SATINALMA KONTROLÜ" bölümü (ADM_SA_CIZ) + düğmeler SAP'DEN OKU
  (AKSORA_SATINALMA_OKU) ve EXCEL'E AKTAR (AKSORA_SATINALMA_EXCEL).
- SAP: CDP oturumu ÜY 7151-7155, EPP 7051-7055. ME5A: StartTransaction, S_WERKS-LOW/HIGH, F8, ızgara cntlGRID1.
  ME2N: LISTU=ALV + Enter, SELPA-LOW=WE101, S_WERKS, F8. Malzeme/talep/sipariş no alanları önce boşaltılır (SET/GET MAT).
  Sütunlar id ya da başlıkla eşlenir (SA_KOLONLAR); tüm sütunlar teşhise yazılır. ME2N açık miktar başlığı "TSL.YAPILACAK".
- Kayıt: _SATINALMA (A1 SA1, B1 okuma zamanı, 3. satırdan T/S, sistem, 16 alan).
- Eşleştirme (SA_ANALIZ): aynı sistem; aynı il serbest (D/K, ATIL_TUM_STOK depolar), atıl miktar (ATIL_HESAPLA x(20)(3)),
  diğer illerden en yüksek serbest stok (transfer). Durum önceliği: ATIL > AYNI İLDE > TRANSFER > STOK YOK.
  Tasarruf = karşılanabilir × AKSORA stok birim değeri (TL). Belge tutarı kendi PB'si ile.
- DOĞRULUK DÜZELTMESİ (03.11 ilk sonuç 291 mn TL şişkindi): aynı stok birden çok kaleme yazılıyordu ve serbest (çalışan)
  stok tasarruf sayılıyordu. Yeni kural: tasarruf YALNIZ atıl stoktan, paylaştırmalı havuz (önce aynı il, sonra diğer illerin
  atılı en büyükten; talepler önce, sonra siparişler); serbest stok yalnız bilgi; stok birimi farklıysa eşleşmez.
  Durumlar: ATIL STOKTAN KARŞILANIR / BAŞKA İLİN ATIL STOĞUNDAN / AYNI İLDE SERBEST STOK VAR / BAŞKA İLDE SERBEST STOK VAR / STOK YOK.
  SAP okurken CoRegisterMessageFilter 0 ile 'OLE eylemi bekleniyor' penceresi bastırılır, sonra eski filtre geri yüklenir.
  ME2N sipariş no boş: liste belgeye göre gruplu; kullanıcı düzende 'Satınalma belgesi' sütununu açıp varsayılan kaydetmeli.
- 2. sonuç (11:20): toplam tasarruf 1.761.896,35 TL (talep 458.963,83 / sipariş 1.302.932,52). CDP ME2N EBELN geliyor
  (/AKSORA düzeni), EPP'de düzen kaydedilmemiş → sipariş no boş. Admin SA düğmeleri xlMove + ADM_GORUNUM'da BX40 satırına hizalanır.
- 3. tur: sipariş no tam (iki sistemde /AKSORA ön ayar). Kontrol: özet = sayfa toplamları, birim farkı 0, atıl taşması yok.
  Eksi 'teslim edilecek' (503… belgeler) → durum FAZLA TESLİM ALINMIŞ (mor, tasarrufa girmez).
  ME2N'de EBELN görünmüyorsa SA_DUZEN_YUKLE: &MB_VARIANT → &LOAD → G51 ızgarasında VARIANT=/AKSORA seçilir
  (başka kullanıcıda ön ayar olmasa da çalışsın; bulunamazsa pencere kapanır, hata yok, yalnız sipariş no boş kalır).
  Admin SA düğmeleri üst il çubuğuna (K4/L4) taşındı: sayfanın altındaki şekiller yakınlaştırmada satır yuvarlamasıyla kayıyordu.
- 4. tur (11:48) Excel doğrulandı: il bazında kalem / tasarruf Özet ve Yönetim Merkezi ile kuruşu kuruşuna aynı, taşma yok.
- MALZEME STOK BİRİMİ SAYFASI: SATINALMA_KONTROL (etiket AKSORA_SATINALMA_V1, M_CDP_RAPOR şablonundan, ATIL_STOK'tan sonra).
  AKSORA_SATINALMA_AC / _IL_SEC / SA_GOSTER / SA_CIZ: KPI (talep, sipariş, atıldan karşılanabilir, tasarruf TL), il düğmeleri,
  SAP'DEN OKU + EXCEL'E AKTAR, il bazında özet (12 sütun), TASARRUF FIRSATLARI (tasarrufa göre azalan, filtreli),
  FAZLA TESLİM ALINMIŞ SİPARİŞLER. Ana panel hızlı erişim: goSatinalma (9 düğme, T_PanelHizala). BIRIM_OLAY / AKTIF_GORUNUM tanır.
- Excel: Özet + Açık Talepler + Açık Siparişler; kırmızı/turuncu/sarı; masaüstü "Satınalma Kontrolü - <tarih>".
- Admin atıl tutarları "#,##0.00 TL" (altın kural).
- DERLEME HATASI (ilk teslim): Private Const SA_KAYIT prosedürlerden sonra yazılmıştı → modülün başına taşındı.
  ME2N sütunları (CDP): EMATN malzeme, WERKS, MENGE, MGLIEF açık miktar (TSL.YPLC.), WTLIEF açık değer, NETPR, SUPERFIELD;
  EBELN gruplandığı için ColumnOrder'da yok → id doğrudan GetCellValue ile denenir; belge boşsa da satır alınır.
  build.sh artık fx/decl_check.py ile 'prosedürden sonra modül bildirimi' denetler (hata varsa durur).

## Satınalma kontrolü düzeltmeleri (03.11.2026)
- SAP'DEN OKU sonrası YÜKLENİCİ_RAPOR'a atlama giderildi: başlangıç sayfası saklanıyor, gizli _SATINALMA oluşturulunca geri etkinleştiriliyor.
- Miktar hücreleri AKSORA_V13_MIKTAR_BICIMI ile: tam sayıda virgül yok, küsuratta virgül (sayfa + Excel çıktısı).
- Excel çıktısında miktarlar 3, tutarlar 2 basamağa yuvarlanıyor (float gürültüsü yok). Özet = sayfa toplamları doğrulandı.

## Satınalma kontrolü tam analiz düzeltmeleri (03.11.2026)
- Doğrulama: 2636 kalem = 81 atıl + 41 fazla teslim + 1246 serbest + 1268 stok yok; atıl paylaştırmada aşım yok; karşılanabilir ≤ açık miktar.
- "Aynı ilde atıl stok" artık ilin toplam atıl stoğu (önceden paylaştırma sonrası kalan gösteriliyordu, aynı malzemede satırdan satıra değişiyordu).
- Sayfa özetinde sayılar binlik noktalı (TOPLAM 1.642 vb.).
- Atıl stok kaynağı / Satıcı metinleri sığdırılıyor (ShrinkToFit).
- Fazla teslim tablosu: "Fazla tutar" pozitif.
- Excel Özet sayfası sayfa özetiyle aynı 11 sütun (toplam tasarruf, fazla teslim, serbest, stok yok), sayılar #,##0.

## Yeni temel up36 + satıcı adı (03.11.2026)
- Temel: kullanıcının 03.11 xlsm'i (up36; yeni Sayfa98 modülü). Kod farkı yalnız harf büyüklüğü.
- Satınalma Excel'i doğrulandı: Özet 11 sütun, toplamlar sayfalarla aynı, aynı il atıl stok değeri sabit.
- Fazla teslim tablosunda satıcı numarası atılıyor (SA_SATICI_AD); ad okunur boyutta.

## Satınalma sütun genişliği (03.11.2026)
- 847319f2 Excel'i bir önceki doğrulanan çıktıyla birebir aynı.
- SA_CIZ: 9. sütun (Atıl stok kaynağı / Satıcı) 125→190 px; Malzeme 140→105, Birim/Belge tarihi/Toplam tasarruf 105→95. Uzun metinler artık küçülmüyor.

## KRİTİK STOK sayfası (03.11.2026, temel up37)
- Malzeme Stok birimi: KRITIK_STOK sayfası (etiket AKSORA_KRITIK_V1, M_CDP_RAPOR kopyası, SATINALMA_KONTROL'ün yanında), ANA PANEL'de 10. hızlı buton goKritikStok (B25:T26), KAPAK açılışında görünür.
- Ek SAP okuması yok. Stok: CDP/EPP/İLÇE/YÜKLENİCİ stok sayfaları (D/K + yüklenici, ATIL ile aynı kapsam). Sarf: _ATIL "Son 6 ay sarf" (MB51). Açık talep/sipariş: _SATINALMA.
- Aylık sarf = 6 ay / 6; kapsam = stok / aylık sarf. KRİTİK <1 ay ve yoldakilerle <3 ay; AZALIYOR <3 ay; YOLDA <3 ay ama yoldakilerle ≥3; YETERLİ 3-12; FAZLA >12.
- DÜZENSİZ SARF: son 6 ayın 3'ünden azında sarf (proje/tek seferlik), kritik sayılmaz. Bunun için _ATIL'e K sütunu "Aylık sarf (6 ay)" eklendi (ATIL_SISTEM_OKU aylık dağılım, ATIL_KAYDET/KAYIT_OKU 11 sütun, kr(8)). Eski kayıtta boş → hepsi düzenli sayılır, sayfada uyarı.
- Önerilen talep = 3 aylık ihtiyaç - stok - açık talep - açık sipariş - diğer illerin atıl stoğundan transfer (aynı sistem, paylaştırmalı, en kritik önce). Tutar = öneri x AKSORA stok birim değeri.
- Excel: Özet / Kritik ve Azalan / Yolda / Tüm İzlenen (22 sütun).
- up37 verisiyle simülasyon (aylık ayrım olmadan): 2980 izlenen; 546 kritik, 251 azalıyor, 46 yolda. Çok sayıda tek seferlik proje sarfı var → aylık ayrım gerekli.

## Kritik stok düzeltmeleri (03.11.2026)
- SATINALMA_KONTROL sekmesi Malzeme birimi açılışında görünür (KAPAK, ATIL/KRITIK gibi).
- Sayfa listesinde "Aylık sarf" yerine "Son 6 ay sarf" (ondalık kafa karıştırıyordu); Excel'de "Aylık ortalama sarf" 2 basamak.
- Aylık dağılım yokken de kesin kural: adetli birimde 6 ayda 3 adetten az sarf → DÜZENSİZ SARF (en fazla 2 ayda olabilir). ÇFT adetli listeye eklendi.
- Kullanıcı Excel'i (fa18117c): 797 kritik+azalan, 75,4 mn TL; yeni kuralla ~608 / 69,5 mn TL. Tutarın çoğu DİREK G-10I / G-K1 (yatırım sarfı); MB51 yeniden okununca aylık ayrım devreye girer.

## İlçe: hızlı buton çıktı düğmeleri, İSG sıralama, aylık demontaj toplam sıralama + Excel (03.11.2026, temel up38)
- AKSORA_YATIRIM_SAYFA_AC ve Y_SAYFA_SEC (olaylar kapalı geçiş) artık AKSORA_CIKTI_DUGME çağırıyor: hızlı butonla gelince ÇIKTI AL / EXCEL'E AKTAR görünür.
- İSG Excel 1. sayfa: I_ISG_SIRALAMA, J:Q sütunlarında ilçe sıralaması (bekleyen İSG değeri az → çok; eşitlikte kalem). Ana tablo A:H değişmedi.
- I_AYLIK_DEMONTAJ: "TÜM AYLAR İLÇE SIRALAMASI" (I_AY_TOPLAM / I_AY_TOPLAM_CIZ): kayıtlı tüm ayların toplam uyumu, fireli ay, fire değeri, en çok fire malzemesi/kategorisi (neden), en kötü ay. İlçe karnesinden sonra.
- EXCEL'E AKTAR (AKSORA_ILCE_AY_AKTAR): Tüm Aylar Sıralama + seçili ayın İlçe Karnesi / Kritik Malzemeler / Fire Listesi. Sayfa aralıkları BX31 karne, BX32 kritik, BX33 fire, BX34 toplam.

## İSG Excel sıralama konumu (03.11.2026)
- Kullanıcı isteği: sıralama ana tablonun altında, aynı A:H sütunlarında (başlık TOPLAM'dan 4 satır aşağıda, sütun başlıkları bir alt satırda). Sütun genişliklerine dokunulmuyor.

## Teknik sağlamlaştırma + TRANSFER LİSTESİ + sade KRİTİK STOK (04.11.2026, temel up40)
- MODÜL BÖLME: ATIL STOK + SATINALMA + KRİTİK STOK + TRANSFER kodu yeni standart modül `modAKSORA_ANALIZ`'e taşındı (BIRIM 697 KB → 494 KB).
  Ortak yardımcılar BIRIM'de Public (Y_*, ADM_*, I_SAYI, I_SIRA_ANAHTAR, BIRIM_MAKRO, ADM_UYAR...), mYonlendirme Public.
  Yeni modül eklemek: tools/add_module.py <bin> <out> <ad> <şablon std modül> (dir + PROJECT + PROJECTwm); bx/xl/vbaProject.bin'e bir kez uygulanır.
  YENİ TEMEL ALINIRKEN: kullanıcının dosyasında modAKSORA_ANALIZ zaten olacak; yoksa add_module tekrar uygulanır.
- build.sh yeni denetimler: fx/xmod_check.py (modüller arası Private erişim = derleme hatası riski), fx/proc_size.py (prosedür ≤48 KB, modül ≤600 KB).
  ADM_KUR 43 KB → 37,8 KB (en değerli depo/malzeme tabloları ADM_EN_DEGERLI'ye ayrıldı).
- Analiz önbelleği: AN_ONB_AC/KAPAT (en fazla 120 sn): ATIL_TUM_STOK ve ATIL_HESAPLA aynı işlemde bir kez hesaplanır (sayfa, Excel, Yönetim Merkezi).
- Veri yaşı: ATIL/SATINALMA/KRİTİK/TRANSFER sayfalarında A9 kırmızı "⚠ ... n GÜN ÖNCE OKUNDU" (AN_ESKI_GUN = 7).
- Yönetim Merkezi: ADM_ANALIZ_UYARI (skor öncesi): MB51 yaşı, satınalma yaşı, fazla teslim, kritik stok, bekleyen transfer kazancı.
- SA_ANALIZ x(24) = atıl kaynakları "ILKODU=miktar;..." ; KS_ANALIZ havuzu satınalmanın ayırdığı atıl stoğu düşer (aynı stok iki kez önerilmez); x(23) transfer kaynakları.
- TRANSFER_LISTESI sayfası (AKSORA_TRANSFER_V1): tek tablo Nereden | Nereye | Malzeme | Kısa metin | Miktar | Birim | Kazanç | Ne yapılacak | Gerekçe; il = gönderen veya alan.
  Excel: Tüm Liste + her gönderen il ayrı sayfa. ANA PANEL 11. hızlı buton goTransfer; KRİTİK ve SATINALMA sayfalarında TRANSFER LİSTESİ düğmesi.
- KRİTİK STOK sadeleşti: 4 kart (kritik, azalıyor, transferle çözülen, satın alınması gereken TL) + tek tablo "SATIN ALINMASI GEREKENLER"; tam liste Excel'de.

## ATIL MB51: pano meşgul hatası (04.11.2026)
- Kullanıcı: EPP 7055 "Pano açılamadı (başka bir program kullanıyor)" → EPP okunmadı. ATIL_PANO_YAZ artık 20 sn bekler (önce 2 sn).
- ATIL_SISTEM_OKU geçici hata verirse aynı sistem bir kez daha denenir (popup kapatılır, 2 sn bekleme).
- Bir sistem bugün okunduysa: "Yalnız <diğeri> okunsun mu?" (Evet/Hayır/İptal) — kısmi okumadan sonra 5 dk tasarruf.

## Satınalma okuma sağlamlığı (04.11.2026, temel up41)
- Kullanıcı: "The control could not be found by id (619)" (ATIL okumasından hemen sonra). Teşhiste bugün ME5A/ME2N satırı yok → seçim ekranında alan bulunamadı.
- SA_TX_AC: işlem açılır, beklenen alan görünene kadar 3 deneme (/n<tx>); açık bilgi penceresi İptal ile kapanır, güvenlik/izin penceresine dokunulmaz.
- SA_YAZ / SA_BAS: alan/düğme yoksa adım + alan adıyla anlaşılır hata ve teşhis satırı.
- SA_SISTEM_OKU: sistem başına okuma; hata olursa bir kez yeniden; yine olmazsa o sistemin ÖNCEKİ kaydı korunur, diğer sistem kaydedilir; hiçbiri okunamazsa kayıt değişmez.
- up41 verisi: _ATIL K (aylık sarf) dolu 3057 kalem. Basit simülasyon (yoldakiler hariç): kritik 128, azalıyor 211, düzensiz 509 (önceden ~546 kritik).

## Kritik/transfer düzeltmeleri (04.11.2026, temel up42; Sayfa100 = TRANSFER_LISTESI)
- Satınalma okuması başarılı (1020 talep / 1633 sipariş). Transfer Excel'i: 119 satır, 1.868.707,24 TL; 94 transfer, 25 kendi atıl stoğu; tekrar yok.
- Adetli malzemede kritik transfer/öneri tam sayı (önceden "1,5 ADT" transfer çıkıyordu).
- Kritik ve transfer sayfasında il yanında sistem (Artvin CDP / Artvin EPP): aynı malzeme iki satır görünüyordu (CDP ve EPP ayrı üretim yeri).

## Malzeme analiz sayfaları birleştirildi + fazla stok kaynağı (04.11.2026)
- Kullanıcı: "çok fazla sayfa, transfer iki yerde; aynı amaca hizmet edenleri tek sayfada birleştir". Karşılaştırılan dış liste: İŞLETME MALZEME DAĞILIM TABLOSU (D001 stok+açık SAS, 2024Q4-2025 sarf payına göre dağıtım).
- Sayfalar artık: ATIL_STOK (kayıt, MB51, kanıt) • SATINALMA_KONTROL = "SATIN ALMA" • TRANSFER_LISTESI.
  - KRITIK_STOK kaldırıldı: SA_GOSTER ilk açılışta etiketli KRITIK_STOK sayfasını siler (AN_ESKI_SAYFA_SIL); AKSORA_KRITIK_AC/IL_SEC/EXCEL → satın alma sayfası.
  - SATIN ALMA sayfası (SA_CIZ yeniden): 4 kart (satın alınması gereken • gereksiz talep/sipariş • transferle çözülen kritik • fazla teslim) + 3 tablo
    (1 satın alınacaklar, 2 iptal/azaltılacak talep-sipariş, 3 fazla teslim), 10 sütun ortak düzen (SA_TABLO_BICIM). İl özeti tablosu kaldırıldı.
  - Satın alma Excel'i: Özet, Açık Talepler, Açık Siparişler + Satın Alınacaklar + Stok Kapsamı (KS_EXCEL_EKLE). Dosya adı "Satın Alma - tarih".
  - ATIL sayfasında transfer görünümü kaldırıldı; 3. sekme TRANSFER LİSTESİ'ne gider (ölü kod: Ayrinti içindeki TRANSFER dalı).
  - ANA PANEL: goKritikStok silinir, goSatinalma "SATIN ALMA"; 10 hızlı buton.
- Fazla stok kaynağı (KS_FAZLA): kritik/azalan ihtiyaç atılla karşılanmazsa diğer ilin serbest (D/K) stoğunun atıl dışı ve 12 aylık ihtiyacını aşan kısmı (adetli: aşağı tam sayı).
  tKay "F:IL=miktar"; AN_KAYNAKLAR → Array(il, miktar, "A"/"F"); transfer satırında "(fazla stoktan)" / "(atıl stoktan)".

## AKSORA imzası + okunur stok süresi (04.11.2026)
- KURAL (kullanıcı): bütün Excel'e aktar çıktılarında AKSORA yazısı silinememeli. T_ExcelBitir başında T_AksoraImza:
  satır 1-2 (başlık + "... | AKSORA ile hazırlanmıştır") kilitli, sayfa şifreli (T_IMZA_SIFRE, TASARIM'da) ama filtre/sıralama/biçim/satır ekle-sil serbest;
  alt bilgi "AKSORA ile hazırlanmıştır"; dosya özellikleri Author/Company = AKSORA, Comments = üretim kaydı; özel özellik AKSORA_IMZA;
  veryHidden "AKSORA_IMZA" sayfası (tarih, kullanıcı, imza no, kaynak dosya).
- Stok süresi ondalık ay yerine metin: AN_SURE_METNI (1,9 ay → "1 ay 27 gün", 0,4 → "12 gün", 0 → "Stok yok", ≥24 ay → yıl). Satın alma sayfası, Excel ve transfer gerekçesi.
- Kullanıcı Excel'i: transfer 256 satır / 7.434.773,60 TL (140 fazla stoktan, 116 atıl/kendi).

## Excel aktar sonrası sayfa atlaması + temel up43 (04.11.2026)
- Kullanıcı: Excel'e aktar sonrası AKSORA'ya dönünce YÜKLENİCİ_RAPOR'a atıyor. Kök neden kesin bulunamadı (son görünür sayfa = gizlenen/eklenen sayfa belirtisi).
  Çözüm: AKSORA_DONUS_KAYDET (her Workbooks.Add öncesi; SA/TR Excel başında zorla) + AKSORA_DONUS_UYGULA (Workbook_Activate ve mUygApp_WindowActivate):
  dönüşte etkin sayfa farklıysa kayıtlı sayfa açılır.
- up43: KRITIK_STOK sayfası silindiği için Sayfa99 modülü kalktı → raw/Sayfa99.txt silindi (yoksa round-trip hatası).
- İmza doğrulandı (kullanıcı Excel'leri): 1-2. satır kilitli, veri satırları açık, AKSORA_IMZA veryHidden, Author=AKSORA. Alt bilgi yazılmamıştı → PrintCommunication açıkken "&8AKSORA ile hazırlanmıştır".

## Temel up44 (04.11.2026)
- Kullanıcı dosyası (7f906351) kod olarak son derlemeyle aynı; temel up44. Excel'lerde alt bilgi XML'de var (&C&8AKSORA ile hazırlanmıştır); openpyxl &R sonrası &C'yi okumuyor, sorun yok.

## ATIL "NE YAPILMALI" + Yönetim Merkezi "BUGÜN YAPILACAK 5 İŞ" (04.11.2026)
- ATIL_STOK varsayılan görünüm ISLEM (ATIL_ISLEM_CIZ): her atıl/atıla yaklaşan kalem için tek eylem satırı:
  1 KENDİ İHTİYACINDA KULLAN (açık talep/sipariş), 2 TRANSFER ET (AN_TRANSFERLER atıl kaynakları; aynı hesap), 3 ÖNCE BUNU KULLAN (ilde tüketiliyor),
  4 BEKLET (başka il kullanıyor), 5 SAT / HURDA / İADE (6 ayda hiçbir ilde sarf yok), 6 ÖNCE KULLAN (3/6 ay kaldı). Sekmeler: NE YAPILMALI | ÖZET | ATIL LİSTESİ | TRANSFER LİSTESİ →.
  ATIL_GOSTER önbellek açar (AN_ONB).
- Yönetim Merkezi: arama bölümünden hemen sonra ADM_YAPILACAK_CIZ: mAdmIsler (ADM_ANALIZ_UYARI doldurur: kritik satın alma, transfer, gereksiz talep, fazla teslim, MB51 yaşı; tutarlı)
  + diğer birimlerin kritik (seviye 1) ve uyarı (seviye 2) maddeleri; ilk 5, No | Yapılacak iş | Tutar | Nerede.

## DERS: VBA büyük/küçük harf duyarsız — saT ve sat aynı ad (04.11.2026)
- "Compile error in hidden module: modAKSORA_ANALIZ": ATIL_ISLEM_CIZ'de "saT As Date" ve "sat As Collection" aynı prosedürde (çift tanım).
  Düzeltme: koleksiyon adı isler. build.sh'a fx/dup_check.py eklendi (aynı prosedürde harf farkıyla çift Dim/parametre; #Else dalı hariç).

## Bugün yapılacak 5 iş: konu başına tek madde (04.11.2026)
- Ekran: 2-5. satırlar Trabzon/Giresun demontaj + Artvin/Trabzon İSG idi (aynı konu tekrar), transfer ve iptal edilecek talep görünmüyordu.
  ADM_YAPILACAK_CIZ: diğer birim uyarılarından her kategoriden (Demontaj, İSG, Teminat...) yalnız en önemli tek madde (önce seviye 1, sonra 2).
- Kullanıcının gönderdiği "Yeni klasör.zip" boş (yalnız boş klasör).

## Tam tur incelemesi (91 ekran görüntüsü, 04.11.2026, temel up45)
- Genel: Satın alma, transfer, atıl NE YAPILMALI, Yönetim Merkezi doğru çalışıyor.
- Tutarsızlık giderildi: ATIL sayfası 4. kart ve Yönetim Merkezi atıl tablosu eski ATIL transfer mantığını (896 kalem / YÜKSEK öncelik) gösteriyordu.
  Artık: ATIL 4. kart "SAT / HURDA / İADE ADAYI" (NE YAPILMALI eylem 5); admin atıl tablosu "Kullanılabilir (kendi / transfer)" ve "Sat / hurda adayı" (ATIL_ISLEM_IL_OZET).
  ATIL_ISLEM_CIZ hesabı ATIL_ISLEM_HESAPLA'ya ayrıldı; eylem satırına 13 il kodu eklendi.
- Bugün yapılacak 5 iş: uzun metin küçülmüyor, satır kaydırmalı (34 px).
- Kullanıcı verisi: ATIL NE YAPILMALI özeti → kendi ihtiyacı 25 / 718 bin TL, transfer 77 / 599 bin, önce bunu kullan 320 / 5,17 mn, beklet 396 / 4,19 mn,
  sat/hurda/iade 735 kalem / 21,1 mn TL, önce kullan 569 / 4,68 mn.

## Katalog v3 + LinkedIn kataloğu (04.11.2026)
- katalog/AKSORA_Katalog_v3.pdf: kurum içi, 19 sayfa, güncel rakamlar (atıl 31,8 mn / transfer 7,4 mn / gereksiz talep 1,6 mn), yeni sayfalar: Atıl NE YAPILMALI, Transfer listesi, Satın alma.
- katalog/AKSORA_LinkedIn.pdf (v2): kurum kataloğu tasarımıyla (A4 yatay) 11 sayfa; kapak aynı, banner'larda kurum logosu ve il haritası bulanık (imgL/). Rakamlar uydurma (atıl 48,6 mn, transfer 12,9 mn, yıllık nakit etkisi 29 mn TL, 2.600 saat/yıl), iller A-E İli, ilçeler İlçe-NN; ekranlar temsilî HTML tablo. Kaynak: scratchpad katalog/linkedin2.html (lhead+lbody).

## İSG listesi + Excel sayım formu (05.11.2026, taban up46)
- İSG: I_ISG_LISTE kullanıcının İSG-HIRDAVAT listesine (05.11) eşitlendi: 372 kalem (283 zimmet / 89 manuel sarf; 7500053643 KABLO SOYMA BIÇAĞI M→Z; 7500051242 MODEM RADYO FREKANS kullanıcı isteğiyle İSG dışı). 80 eksik kalem eklendi, 7500053995 (ÇANTA MALZEME EL TİPİ) M→Z düzeltildi. docs/isg/ISG_ISLEM.csv aynı liste. İSG sayfaları/mantığı değişmedi. Not: listedeki kalemler demontaj kontrolüne girmez (mevcut kural).
- Sayım: stok sayfalarında ÇIKTI AL → "EVET = Excel / HAYIR = PDF". Excel sayım formu (T_SayimExcelYaz): PDF ile aynı düzen (A4 yatay, başlık, kaşe/imza bandı, sayfa no, siyah-beyaz baskı), SAYILAN (sarı, ≥0 doğrulama) ve AÇIKLAMA doldurulur, FARK formülle (eksi kırmızı, artı yeşil), 3. satırda "Girilen x / n, Farklı çıkan y". Stok sütunları ve FARK kilitli (T_AksoraImza + mSayimKilit). Masaüstüne "<depo> - SAYIM - <sayfa> - <tarih>.xlsx". Koşullu biçimde formül yok (Türkçe Excel yerel formül sorunu).

- İSG bölge sarfı (05.11): manuel sarf kaleminde birim fiyat (stok değeri / miktar) > 10.000 TL ise "BÖLGE SARF EDECEK" (gri). Sabit I_ISG_BOLGE_TL (modAKSORA_BIRIM başı). İlçe özetinde ayrı "Bölge Sarfı Kalem" sütunu; ilçe durumuna, MANUEL SARF sayılarına, Yönetim Merkezi İSG uyarı/karnesine girmez. 4. filtre düğmesi BÖLGE SARF EDECEK; Excel aktarda gri renk.

- İSG bölge sarfı görünümü (05.11): ilçeye iş yok ama bölge kalemi varsa Durum "BÖLGE SARFI BEKLİYOR" (gri); Excel aktar 1. sayfada "Bölge Sarfı Kalem" sütunu; sıralamada bu ilçeler temizlerle önde, neden "İlçeye iş yok: değer bölge sarfı". Manuel sarf miktarı (adet toplamı) kullanıcı isteğiyle aynı kaldı.

## HATA DERSİ (06.11.2026, taban up47): İSG'ye bütün malzemeler geldi
- Sebep: ADM_VERI'de `If liste.Exists(mal) And ... liste(mal) ...` — VBA And kısa devre yapmaz; liste(mal) okuması olmayan anahtarı Static sözlüğe ekledi (değer Empty → "MANUEL SARF"). Düzeltme: önce Exists, sonra oku (isgVar). I_ISG_LISTE sayı tutmazsa listeyi yeniden kurar.
- KURAL: Dictionary'de d(k) okuması anahtar EKLER; Exists ile aynı satırda And/Or içinde asla okuma yapma.

- İSG bölge sarfı sınırı (06.11, taban up48): 12.000 TL ve üzeri (>=, 0,005 tolerans). Sabit I_ISG_BOLGE_TL = 12000.

## Stok karşılaştırma: NEREDEN / NEREYE (06.11.2026, taban up49)
- STOK_KARŞILAŞTIRMA tablosuna N sütunu "NEREDEN / NEREYE" (STOK_KARS_NEREDEN). SAP yok; 5 ilin stok hafızasından aynı malzeme+birimin azalış↔artış eşleştirmesi. Aynı il: birebir miktar, sonra büyükten küçüğe dağıtım. İller arası: yalnız birebir aynı miktar ve iki yönde tek aday. Karşı iller AKFIX_KARS_IL_DURUM(il, t+1 saat) (kayıt saati farkı). Karşılıksız: "Sarf / depo dışı" / "Alım / depo dışı"; HURDA partisi "Hurda girişi (demontaj)" / "hurda çıkışı (satış / imha)", ÇIKMA "Çıkma girişi (demontaj)" / "çıkma malzeme kullanıldı". Aynı depoda parti değişimi partiyle yazılır.
- Eski sayfada ilk karşılaştırmada STOK_KARS_NN_HAZIRLA: N genişler (44), sağ panel aynı miktar kayar, başlık A1:N, filtre A15:N.
- Simülasyon (fx/sim_nn.py) Artvin 05.11→06.11: 7055 D002 → 7055 D011 traversler; Trabzon 7052 D001 ← 7053 D014 hücre (Gümüşhane tarafı da →).

- NEREDEN/NEREYE ek (06.11): karşı taraf ilçe (K) ve yüklenici (M) depoları: _KH_<IL> günlük kayıt (STOK_KH_KAYDET; toplu güncellemede _AKLIVE_ ILC+YUK, tek il güncellemesinde STOK_KH_KAYDET_CANLI; 60 gün, gün başına son kayıt). _KH_ yoksa ilçe biriminin _IH_ geçmişi (yalnız K). Gün kuralı: t günü sütunu, yoksa t öncesi son. Simülasyon: 7372 D001 -9 → 7155 K009.
- Üst bantlar ve hızlı düğmeler N'ye kadar (STOK_KARS_UST_GENISLET, eski sayfa için; yeni kurulum doğrudan N).

- İÇ NAKİL (06.11): aynı üretim yerinin D depoları arasında tamamı eşleşen satır durumu "İÇ NAKİL" (gri); azalan/artan sayı ve filtrelerine girmez.
- Stok karşılaştırma sağ paneline BÜTÜN İLLERİ GÜNCELLE (AKSORA_KARS_TOPLU_DUGME, AKSORA_V11_KARS_UYARI sonunda; makro AKSORA_KARS_BUTUN_ILLER: 5 il güncelle → karşılaştırma sayfası → tarihler → KARŞILAŞTIR). Taban up50.

## HATA DERSİ (06.11.2026): "Rapor kaynağı boş: STOKLAR" / "Stok sayfası boş"
- Sebep: STOKLAR'da kullanıcı filtresi açık (D003/D009/D010 + İSG/HIRDAVAT); güncellemeden sonra AKG_KOPYALA filtreyi yeniden uyguluyor, Range.Find filtre gizli satırları atlıyor → AKSORA_SON_SATIR başlık satırını döndürüyor.
- Düzeltme: AKSORA_SON_SATIR FilterMode'da UsedRange sütununu doğrudan okur (filtreye dokunmaz). KURAL: son satır hesabı asla yalnız Find'a bırakılmaz.

## Ortak klasöre YAYIN (09.11.2026)
- Bütün illeri güncelle ve Yönetim Merkezi Tüm birimleri güncelle sonunda AKSORA_YAYINLA_SESSIZ: ortak klasöre ThisWorkbook.SaveCopyAs "AKSORA_YAYIN.xlsm" (salt okunur özniteliği), açıksa "AKSORA_YAYIN_yyyymmdd_hhnn.xlsm"; AKSORA_YAYIN_BILGI.txt. Klasör _AKSORA_YAYIN_KLASOR (ilk sefer sorulur; "YOK" = kapalı). Alt+F8: AKSORA_YAYIN_AYAR, AKSORA_YAYINLA.
- İzleme modu = dosya adında AKSORA_YAYIN (AKSORA_YAYIN_MODU): SAP_SESSION_AL, AKSORA_SAP_GUNCELLEME_IZNI, BÜTÜN İLLER, TÜM BİRİMLER kapalı (AKSORA_YAYIN_UYARI). Sayfalar, filtre, Excel'e aktar, çıktı, birim şifreleri aynen.
- Kullanıcıya önerilen yol: kişisel SAP kullanıcısı yerine IT'den salt okunur teknik SAP kullanıcısı + sunucuda zamanlanmış çalışma (faz 2), Power BI (faz 3).

## 11.11.2026 (taban up51)
- Kapanışta çift "kaydedilsin mi": Workbook_BeforeClose sonunda ThisWorkbook.Saved = True (UYGULAMA_MODU_KAPAT dosyayı kirletiyordu).
- Müşteri stok filtresinde sayfa atlaması: AKSORA_MUSTERI_GECIS_KORU artık her hedef sayfa için (2 sn içinde) stok sayfasına döndürür (önce yalnız M_MALZEME_ARA).
- Teminat Konum: her hücrede sağda ▼ (NumberFormat "@* "▼"", şekil yok); tıklayınca CommandBar açılır menü (AKSORA_KONUM; AKSORA_YATIRIM_KONUM_SECIM → AKSORA_YATIRIM_KONUM_SEC, satır bellekte mKonumSatir — hücreye yazmak her tıklamada tüm kitabı yeniden hesaplatıp titretiyordu; ONAR hazırsa atlar; menü bir kez kurulup yeniden kullanılır; menü kapanınca seçim aynı satır F sütununa alınır ki aynı Konum hücresine tekrar tıklanabilsin). Form kutusu ve doğrulama listesi kaldırıldı. T_CiktiListe '@*' biçimini '@' yapar. Değişimden sonra Y_TEMINAT_KAL.
- Teminat Konum listesi: Y_KONUM_LISTE sabit "Depoda,Sahada,Kısmi"; Y_TEMINAT_LISTE_ONAR sayfa açılışında (AKSORA_BIRIM_OLAY) listeyi ve kilit açıklığını yeniden uygular.

## PROJE MALZEME TAKİBİ (11.11.2026) — Y_HAKEDİŞ sayfası yerine
- Kaynak: kullanıcının manuel Excel'i (PROJE ENVANTER keşif matrisi + SARF-2510 MB51 221 + D002/GR-1/GR-2/KET MB52). SAP akışı: 101 → HAVUZ.7050 (Q proje stoğu), 415 HAVUZ → proje PYP, 221 sarf (belge başlığı N.HAKEDİŞ). MB51 sütunu "Proje yapı planı öğesi" = satırın kendi PYP'si, "PYP öğesi" = karşı.
- Doğrulama: tüm hareketli MB51'den yüklenici depo stokları 23.10 MB52 ile M122 75/75, M123 111/111, M124 89/89; malzeme İHTİYAÇ kullanıcı BO sütunu 167/167 (fx/prj_sim.py).
- SAP: MB51 düzeni /AKSORA_PRJ (16 sütun, düz liste). AKSORA_PRJ_MB51_OKU (modAKSORA_ANALIZ): EPP ÜY, depo D002 + AKSORA_IL_YUKLENICI (Z hariç), yıl başı–bugün, hareket türü boş. Kayıt _Y_PRJ_<IL> (A1 PRJ_V1).
- Keşif: AKSORA_YATIRIM_KESIF_YUKLE — matris (PYP satırı + ad satırı + malzeme sütunu) ya da tek proje (Y_PROJE_TARA + InputBox PYP). Kayıt _Y_KESIF (il,PYP,mal,miktar,tanım; "#il" satırı proje adı/dosya/tarih). Yüklenen projelerin eski keşfi değişir.
- Hesap (Y_PRJ_VERI): proje stoğu = PYP hareket toplamı; havuz/serbest = MB52 depo stoğu − projelerde kalan; fiyat MB52 değer/miktar, yoksa MB51 tutar/miktar. Sayfa: kartlar, YÜKLENİCİ DURUMU, PROJELER (Durum EKSİK/FAZLA/TAMAM/BAŞLAMADI), MALZEME DETAYI (seçili proje: hakediş sütunları, en çok 6; tümü: İhtiyaç = keşif − sarf − projelerde − yüklenici havuzu − D002). Sürüm BX25=PRJ_V1. Panel düğmesi "PROJE / HAKEDİŞ".

- 2026-11-11: Proje takibi düzeltme: boş veride IIf içindeki 0/0 "Overflow" verdi (IIf iki dalı da hesaplar); MB51 kısa başlıklar (Kayıt trh., Blg.bşlk.metni, Hrk.tü.mtn, Ö, iki "PYP öğesi") eşlendi — iki PYP sütununda daha dolu olan satırın kendi PYP'si.

- 2026-11-11: Proje takibi + açık siparişler: SAP'DEN OKU, MB51'den sonra ME2N (WE101, ilin EPP ü.yeri) okur → _Y_SAS_<IL> (belge, kalem, depo, mal, tanım, açık mkt, birim, teslim, belge trh, satıcı, net fiyat, PB). Yalnız D002 satırları ihtiyaçtan düşülür; diğer depolar not. Tüm projeler tablosu: "Açık Sipariş (D002)" sütunu, Ne Yapılmalı: Satın al / Siparişte (teslim, GECİKMİŞ) / Sipariş gereksiz (iptal-azalt) / Fazla / Tamam; aynı tanımla başka kodla sipariş uyarısı (örn. 1000005087 ↔ 1000009010); fiyatsız malzemeye SAS TL net fiyatı. Sayfa sürümü BX25=PRJ_V2. Taban up52.

## Sıradaki işler
- (Canlı güncelleme şeridi tüm birimlere eklendi, yukarıya bakınız.)
- ATIL kuralı 2 ilk tam okuma (CDP+EPP, 7 dk 52 sn): ATIL 1.509 kalem / 31,35 mn TL; 3 AY 319 / 3,31 mn; 6 AY 368 / 1,84 mn; transfer 906. (Eski kural: 2.435 / 637,6 mn.) Doğruluk testi kullanıcıda.
- Kalite Kontrol birimi: kullanıcıdan MB51 /KK düzeni, IQ09 üretici/garanti alanları ve tasarım onayı bekleniyor.
