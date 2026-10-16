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

## Sıradaki işler
- (Canlı güncelleme şeridi tüm birimlere eklendi, yukarıya bakınız.)
