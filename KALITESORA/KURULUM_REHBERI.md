# KALİTESORA v2.3 — Kurulum ve Kullanım

**Çoruh Bölgesi malzeme kalite, garanti ve arıza takip merkezi.** Tamamen VBA ile çalışır; Excel dışında bir şey gerektirmez.
Tasarım ailesi AKSORA ile aynıdır: banner görseli, açık zemin, gölgeli beyaz kartlar, hizalı sekme çubuğu.

> **Not:** Kod bu ortamda Excel'de çalıştırılamadı. Statik denetimden geçti: blok eşleşmesi, tanımsız
> değişken, anahtar kelime çakışması, yordam ortasında tanım, parametre sayısı ve düğme makro adları.
> Bir hata çıkarsa mesajdaki **Adım** ve **Hata** satırlarının ekran görüntüsünü gönderin.

---

## 1. Kurulum (hazır paket)

1. `KALITESORA_PAKET.zip` dosyasını bir klasöre açın. Klasörde şunlar olmalı:
   ```
   KALITESORA.xlsm
   KALITESORA_GORSEL\   (banner, zemin, logo)
   ```
   Görsel klasörü dosyanın **yanında** durmalıdır; kurulum görselleri buradan okuyup dosyaya gömer.
2. `KALITESORA.xlsm` > sağ tık > **Özellikler** > **Engellemeyi kaldır** > Tamam.
3. Dosyayı açın > **İçeriği Etkinleştir**. Kurulum kendiliğinden başlar; **Evet** deyin.
4. Kurulum bitince dosyayı kaydedin. Bundan sonra görsel klasörü gerekmez.

Daha önce v1 kuruluysa: v2'de tablo yapısı değişti. Yeni paketi **boş** olarak kurun; v1'de veri girdiyseniz
haber verin, aktarma yolunu birlikte yapalım.

---

## 1a. Uygulama modu ve ekran doldurma (v2.1)

- Dosya açılınca **uygulama modu** devreye girer: Excel şeridi, formül çubuğu, satır/sütun başlıkları ve
  kılavuz çizgileri gizlenir. Başka bir Excel dosyasına geçince ya da dosya kapanınca Excel görünümü geri gelir.
- Her sayfa açıldığında ve pencere boyutu değiştiğinde ekran kendiliğinden doldurulur: banner ve sekme bandı
  pencere genişliğini tam kaplayacak şekilde yakınlaştırma ayarlanır.
- Excel görünümüne geçmek için: PANEL > **EXCEL / UYGULAMA** düğmesi (YÖNETİM sayfasında da var).
- v2.0 kurulu dosyada v2.1 ilk açıldığında **tasarım güncellemesi kendiliğinden** uygulanır; kayıtlar korunur,
  önce yedek alınır. Bu sırada `KALITESORA_GORSEL` klasörü dosyanın yanında olmalı; klasör yoksa mevcut banner korunur.
- Tablo başlıkları tek renk (lacivert). Elle yazılabilen sütunlar her sayfanın açıklama satırında belirtilir.

## 2. Sayfalar

| Sekme | İçerik |
|---|---|
| **PANEL** | 4 özet kartı (sevkiyat, arıza, garanti, üretici dönüşü) ve 4 uyarı kartı: garantisi bitecek, teslimatı geciken, üreticide bekleyen, garantisi girilmemiş. **Karta tıklayınca ilgili liste süzülmüş açılır.** Altta sevk yerine göre özet ve durum özeti (kalem / tutar / pay) tabloları. |
| **KAYIT** | Sekmeli tek form: 01 MALZEME · 02 SEVKİYAT · 03 ARIZA. **Kod ya da ad yazınca hafızadaki her şey dolar**; sağdaki HAFIZA kartında geçmiş görünür. Kayıtlı numara yazılırsa kayıt forma gelir; KAYDET yeni kayıt ekler ya da günceller. |
| **SEVKİYATLAR** | **YILLIK EXCEL YÜKLE** (SAP / ÖRNEK.xlsx biçimi), hazır süzgeçler. Sarı başlıklı sütunlar (Gerçek Geliş, **Garanti (Ay)**, Belge No, Seri, Not) doğrudan düzenlenir; hesaplar kendiliğinden güncellenir. |
| **ARIZALAR** | Tüm arızalar; süzgeçler (devam eden, üreticide, çözülen, garanti kapsamında). Mail dosyası sütunu tıklanabilir bağlantıdır. |
| **MALZEMELER** | Sade ana liste. **Garanti (Ay)** buradan girilirse garantisi boş tüm sevkiyatlarına uygulanır. |
| **KARNE** | Tek arama kutusu: malzeme kodu / adı **ya da** üretici kodu / adı. Uygun karne kendiliğinden açılır. Ayrıca **TEDARİKÇİ SIRALAMASI** (kalite skoru A-D). |
| **RAPORLAR** | 5 yıllık karşılaştırma, aylık dağılım, 2 grafik. |
| **YÖNETİM** | TANIMLAR (sevkçi, sevk yeri, kategori, **döviz kurları**, parametreler), TOPLU AKTAR, İŞLEM GÜNLÜĞÜ ve yönetim araçları. |

---

## 3. Arıza kaydı nasıl akar?

1. **Arıza Tarihi**, **Malzeme Kodu** ya da **Malzeme Adı** girilir. Üretici, bağlı sevkiyat, sevk yeri ve garanti bitişi dolar.
2. **Seri No** girilirse aynı seri numaralı sevkiyat seçilir. Seri yoksa arıza tarihinden önceki en son teslim seçilir.
3. **Garanti devam ediyor mu?** otomatik EVET/HAYIR gelir. Elle değiştirilirse "ELLE" olarak saklanır.
4. **Arıza Sebebi / Açıklama**, **Üretici ile Mail Yazışması** (yapıştırılır) ve **DOSYA EKLE** girilir.
   DOSYA EKLE ile .msg / .eml / .pdf eklenir; kayıtta `KALITESORA_MAILLER\ARZ-xxxxxx\` klasörüne kopyalanır.
5. **Arıza devam ediyor mu?** ve **Çözüldü mü?** (EVET / HAYIR / KISMEN) seçilir. Durum buna göre hesaplanır:
   ÇÖZÜLDÜ · KISMEN ÇÖZÜLDÜ · ÇÖZÜLEMEDİ · ÜRETİCİDE · DEVAM EDİYOR.

---

## 4. Yıllık Excel (SAP) yükleme

SEVKİYATLAR > **YILLIK EXCEL YÜKLE** > dosya > başlangıç tarihi.

- Başlıklar adla bulunur (ÖRNEK.xlsx'teki gibi; *Tedarikçi adı* ve *Şirketin adı* sütunları da desteklenir).
- TANIMLAR'da olmayan üretim yerleri bulunursa sorulur: eklensin mi, atlansın mı.
- Aynı SAS belge + kalem + malzeme ikinci kez aktarılmaz; aktarım tekrar çalıştırılabilir.
- Döviz fiyatlar olduğu gibi saklanır; TL tutar **YÖNETİM > Kur (TL)** tablosuyla hesaplanır. Kur girilmezse tutar boş kalır.
- SAP'ta mal giriş tarihi ve garanti süresi yok:
  - **Geliş tarihi** için aktarımda sorulur (tamamen teslim alınan kalemlerde SAP teslimat tarihi, yaklaşık olarak).
  - **Garanti süresi** için MALZEMELER > **GARANTİSİ GİRİLMEMİŞ** süzgeciyle her malzemeye bir kez girin. Aynı malzemenin boş garantili tüm sevkiyatlarına uygulanır.

---

## 5. Saklama, yedek, güvenlik

- Hiçbir makro gerçek kaydı kendiliğinden silmez (yalnız **ÖRNEK VERİ** notlu deneme satırları silinebilir).
- Kurulum, aktarım ve SAP yüklemesi öncesinde `KALITESORA_YEDEK` klasörüne otomatik yedek alınır. Son yedek 7 günü geçerse açılışta sorulur.
- Excel tablosu ~1 milyon satır taşır; 5+ yıl için yeterlidir.
- Sayfa koruma şifresi: **`KS2026`** (YÖNETİM > Sayfa koruması aç / kapat).

---

## 6. Elle kurulum (gerekirse)

1. Boş bir `.xlsm` dosyasının yanına `KALITESORA_GORSEL` klasörünü koyun.
2. Güven Merkezi > Makro Ayarları > "VBA proje nesne modeline erişime güven" seçeneğini açın.
3. Alt+F11 ile `vba\modKS_YUKLEYICI.bas` dosyasını içe aktarın.
4. Alt+F8 ile `KALITESORA_YUKLE` makrosunu çalıştırın.

`.bas` dosyaları Windows-1254 (Türkçe) kodlamasındadır.

## v2.3 yenilikleri
- Açılışta tasarım güncellemesi kendiliğinden uygulanır (önce yedek alınır, kayıtlar korunur).
- **ARA** düğmesi: SEVKİYATLAR / ARIZALAR / MALZEMELER tablolarında kod, ad, seri veya üreticiye göre hızlı süzme.
- **PDF**: RAPORLAR ve KARNE sayfaları tek tıkla `KALITESORA_RAPOR` klasörüne yatay PDF olarak kaydedilir.
- **ÜRETİCİYE MAİL**: KAYIT > Arıza ekranından Outlook taslağı hazırlanır (Outlook yoksa metin alana yazılır); bildirim tarihi boşsa bugün yazılır.
- Mail dosyası **AÇ** düğmesi; eklenen mail/sohbet dosyası doğrudan açılır.
- Tarih ve sayı sütunlarına giriş denetimi; hesaplanan sütunlar ve tablo altı kilitli.
