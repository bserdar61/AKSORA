# KALİTESORA — Kurulum ve Kullanım Rehberi

**Çoruh Bölgesi malzeme kalite, garanti ve arıza takibi.** Tamamen VBA; Excel dışında bir şey gerektirmez.
Tasarım ailesi AKSORA ile aynıdır: lacivert bant, camgöbeği vurgu, beyaz kartlar, Segoe UI.

> **Dürüst not:** Kod bu ortamda Excel ile derlenip çalıştırılamadı (Excel yok). Statik denetimden
> geçti: blok eşleşmeleri, tanımsız değişken (Option Explicit), tanımsız yordam, parametre sayısı,
> yinelenen tanım. İlk kurulumda çıkan her hata mesajını (Adım + Hata satırı) olduğu gibi iletin;
> birlikte düzeltiriz.

---

## 1. Kurulum (yaklaşık 2 dakika)

1. Excel'de **boş bir çalışma kitabı** açın ve **Makro İçerebilen Excel Çalışma Kitabı (.xlsm)** olarak
   kaydedin (örn. `KALITESORA.xlsm`). Dosyayı OneDrive adresinden değil, bilgisayardaki bir klasörden açın.
2. Bir defalık ayar: **Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarları > Makro Ayarları >
   "VBA proje nesne modeline erişime güven"** kutusunu işaretleyin.
3. `vba` klasörünü (ya da `KALITESORA_VBA.zip` içeriğini) bilgisayara indirin.
4. **Alt+F11** > **File > Import File…** > yalnızca **`modKS_YUKLEYICI.bas`** dosyasını seçin.
5. **Alt+F8** > `KALITESORA_YUKLE` > Çalıştır > `.bas` dosyalarının bulunduğu klasörü seçin.
   Yükleyici 12 modülü içe aktarır, açılış kodunu ekler ve kurulumu başlatır.
6. Dosyayı kaydedin.

**Elle kurulum (2. adımdaki ayar açılamıyorsa):** 12 `modKS_*.bas` dosyasının hepsini **File > Import File**
ile tek tek içe aktarın. `ThisWorkbook_KODU.txt` içindeki 3 satırı ThisWorkbook'a yapıştırın.
Ardından **Alt+F8 > `KALITESORA_KUR`** makrosunu çalıştırın.

> `.bas` dosyaları **Windows-1254 (Türkçe)** kodlamasındadır. Türkçe Windows'ta karakterler doğru gelir.
> Başka dilde bir Windows'ta Türkçe harfler bozuk görünebilir.

`KALITESORA_KUR` makrosu tekrar çalıştırılabilir (**onarım modu**). Bu modda önce yedek alınır, veri
tablolarındaki kayıtlara dokunulmaz, form, panel ve rapor tasarımı yeniden çizilir.

---

## 2. Sayfalar ve aralarındaki ilişki

```
TANIMLAR ──► açılır listeler (sevkçi, sevk yeri, kategori, arıza tipi …)
   │
MALZEMELER (Malzeme Kodu = anahtar)
   │  1 ─── n
SEVKİYATLAR (Sevk No = SVK-000001)   ← Malzeme Kodu ile bağlı; garanti süresi malzemeden gelir
   │  1 ─── n
ARIZALAR (Arıza No = ARZ-000001)     ← Sevk No ile bağlı; malzeme, sevkçi, yer, garanti bitişi sevkiyattan gelir
   │
   ├─► PANEL            10 KPI kartı, 4 canlı liste, 2 grafik
   ├─► MALZEME KARNESİ  tek malzemenin tüm sevkiyat + arıza geçmişi
   ├─► PERFORMANS       sevkçi / üretici kalite skoru (A-B-C-D)
   ├─► 5 YILLIK RAPOR   yıl yıl karşılaştırma + aylık dağılım
   └─► GÜNLÜK           her işlemin izi (kim, ne zaman, ne yaptı)
```

| Sayfa | Ne işe yarar |
|---|---|
| **PANEL** | Dönem seçimi (TÜMÜ / yıl). Kartlar: sevk kalemi, sevk tutarı, zamanında teslim %, bekleyen teslimat, garantisi bitenler, arıza kaydı, garanti içi arıza %, açık arıza, ort. üretici dönüşü, tamir maliyeti. Listeler: en çok arıza veren malzemeler, açık / üreticide arızalar, garantisi bitecekler, gecikenler. |
| **KAYIT** | Üç kartlı giriş formu. Her kartta KAYDET / GÜNCELLE / GETİR / TEMİZLE düğmeleri var. `»` işaretli alanlar otomatik dolar. |
| **SEVKİYATLAR / ARIZALAR / MALZEMELER** | Ana veri tabloları. Filtrelenebilir. Turkuaz başlıklı sütunları sistem hesaplar. Düzeltmeler KAYIT sayfasından GETİR + GÜNCELLE ile yapılır. |
| **TOPLU AKTAR** | Geçmiş yılların verisini yapıştırıp form kurallarıyla doğrulayarak aktarır. **SAP LİSTESİNDEN AKTAR** düğmesi SAP satınalma listesini doğrudan okur. |
| **TANIMLAR** | Açılır liste kaynakları, parametreler ve **ARAÇLAR** düğmeleri. Bu sayfa korumasızdır, tablolara satır ekleyebilirsiniz. |

---

## 3. Otomatik hesaplanan alanlar

| Alan | Kural |
|---|---|
| Toplam Tutar | Adet × Birim Fiyat |
| Gecikme (Gün) | Gerçek Geliş − Planlanan Teslim. Geliş yoksa ve tarih geçtiyse bugüne kadarki gecikme yazılır. |
| Teslim Durumu | ZAMANINDA / GECİKMELİ / GECİKİYOR / BEKLENİYOR |
| Garanti Bitiş | **Gerçek Geliş + Garanti (Ay)**. Garanti süresi sevk anında malzemeden alınır. Malzemenin garantisi sonradan değişirse eski sevkiyatlar etkilenmez. |
| Garanti Durumu | DEVAM / BİTİYOR (n GÜN) / BİTTİ. Uyarı süresi TANIMLAR > `Garanti bitiş uyarısı` (varsayılan 60 gün). |
| Garanti Kapsamı | Arıza Tarihi ≤ Garanti Bitiş ise EVET, değilse HAYIR. Geliş tarihi yoksa BELİRSİZ. |
| Dönüş Süresi | Üretici Dönüş − Üreticiye Bildirim (gün) |
| Çözüm Süresi | Çözüm Tarihi − Arıza Tarihi (gün) |
| Durum | AÇIK → ÜRETİCİDE (bildirim var) → DÖNÜŞ ALINDI → KAPANDI (çözüm tarihi var) |
| Önceki Arıza Tarihi / Tekrar Sayısı | Aynı **seri no**'lu malzemenin, seri no yoksa aynı **sevkiyatın** daha önceki arızası |
| Malzeme: Arıza Oranı | Toplam arızalı miktar / toplam sevk miktarı |

**Doğrulama kuralları:** Geliş, arıza, bildirim, dönüş ve çözüm tarihleri bugünden ileri olamaz. Arıza tarihi
geliş tarihinden önce olamaz. Dönüş tarihi bildirimden önce olamaz. Arızalı adet sevk edilen miktarı aşamaz.
Aynı irsaliye + malzeme ikinci kez girilirse uyarı çıkar.

**Kalite skoru (PERFORMANS):**
`100 − en çok 50 (arıza oranı % × 5) − en çok 25 (ort. gecikme gün × 1) − en çok 25 (ort. dönüş gün × 0,5)`.
Notlar: A ≥ 85, B ≥ 70, C ≥ 50, D < 50. Katsayılar TANIMLAR sayfasından değiştirilebilir.

---

## 4. 5 yıl ve sonrası: saklama ve yedek

- Hiçbir makro kayıtları kendiliğinden silmez. Tek istisna, yalnız **ÖRNEK VERİ** notlu deneme satırlarını
  silen `ÖRNEK VERİYİ SİL` düğmesidir.
- Excel tablosu ~1 milyon satır taşır. Yılda 20.000 sevkiyat + arıza olsa bile 5 yıl ≈ 100.000 satırdır.
  Hesaplar dizi (array) üzerinde yapıldığı için hızlı kalır.
- **YEDEK AL** düğmesi dosyanın kopyasını yanındaki `KALITESORA_YEDEK` klasörüne tarih-saat adıyla koyar.
  Kurulum, toplu aktarım, SAP aktarımı ve örnek veri silme öncesinde yedek kendiliğinden alınır.
- Dosya açılırken son yedek 7 günden eskiyse (ayarlanabilir) yedek alınıp alınmayacağı sorulur.
- Sayfa koruma şifresi: **`KS2026`**. ARAÇLAR > *Sayfa korumasını aç / kapat*.

---

## 5. Tedarikçi listesi (hafızada)

`modKS_TEDARIKCI` modülü, Çoruh bölgesine **01.01.2026 ve sonrası** teslimat tarihli SAP kalemlerindeki
**47 tedarikçi kodunu** içerir. Kaynak: `TEDARİKÇİ İSİM.xlsx`, `cdp` sayfası. Kurulumda TANIMLAR'a yazılır.
Okunabilir kopya: `veri/tedarikciler_2026.csv`.

- Sevkçi değeri SAP biçimindedir (`1300001416 LUNA ELEKTRİK ELEKTRONİK`). Böylece aynı ad farklı kodla
  karışmaz ve SAP aktarımı doğrudan eşleşir.
- "Tür" sütunu SAP malzeme kodundan yapılmış bir tahmindir: 3xxx hizmet, 1xxx ve 7xxx malzeme.
- 1200007200 FIRAT ELEKTRİK DAĞITIM bir grup şirketidir ("Kurum İçi" olarak işaretli).
- Aynı firma farklı kodla da geçer: REFORM ALTYAPI, İŞTEKS, VERİ BİLGİ MERKEZİ.
- ARAÇLAR > **TEDARİKÇİ LİSTESİNİ YÜKLE** yalnız eksik kodları ekler, mevcutları değiştirmez.

---

## 6. SAP listesinden aktarım

TOPLU AKTAR > **SAP LİSTESİNDEN AKTAR** > SAP dosyasını seçin > başlangıç tarihini girin.

- Başlıklar adla bulunur: *Satıcı/teslimat yapan ürt.yeri, Malzeme, Kısa metin, Teslimat tarihi,
  Üretim Yeri Adı, Miktar (depo ÖB), Girilen miktar, Net fiyat, Para birimi, Satınalma belgesi, Kalem*.
- Yalnızca TANIMLAR'daki sevk yerlerine (ÇEDAŞ 7151–7155) ait satırlar alınır. FEDAŞ gibi bölge dışı satırlar atlanır.
- Kaydı olmayan malzemeler için malzeme kartı açılır. Garanti süresi varsayılan 24 aydır
  (TANIMLAR > parametre). **Bu süreyi sözleşmeye göre düzeltin.**
- İrsaliye alanına `SAS belge/kalem` yazılır. Aynı kalem ikinci kez aktarılmaz, aktarım tekrar çalıştırılabilir.
- **SAP listesinde mal giriş tarihi yoktur.** Aktarım sırasında sorulur: tamamı teslim alınmış kalemlerde
  geliş tarihi olarak SAP teslimat tarihi yazılsın mı? EVET seçilirse notta "yaklaşık" yazar.
- Döviz (USD/EUR) fiyatlı kalemlerde TL tutarı **0** yazılır, gerçek fiyat ve para birimi Not'a eklenir.
  Kur dönüşümü yapılmaz.

---

## 7. Makro listesi

| Makro | Ne yapar |
|---|---|
| `KALITESORA_YUKLE` | Modülleri içe aktarır ve kurar (ilk kurulum) |
| `KALITESORA_KUR` | Kurulum / onarım |
| `KS_PANEL_YENILE` | Paneli ve 5 yıllık raporu yeniler |
| `KS_YENIDEN_HESAPLA` | Tüm türetilmiş alanları baştan hesaplar |
| `KS_YEDEK_AL` | Elle yedek alır |
| `KS_SAP_AKTAR` | SAP satınalma listesini aktarır |
| `KS_ORNEK_VERI_YUKLE` / `KS_ORNEK_VERI_SIL` | Kurgusal deneme verisi (yalnız tablolar boşken yüklenir) |
| `KS_TEDARIKCI_YUKLE` | 2026 tedarikçi listesini TANIMLAR'a ekler |
| `KS_KORUMA_DEGISTIR` | Sayfa korumasını açar / kapatır |
| `KS_VBA_DISARI_AKTAR` | Kodu masaüstüne dışa aktarır (bana göndermek için) |
