# AKSORA — İnceleme Raporu ve v3.1 Geliştirmeleri

İncelenen sürüm: `AKSORA.xlsm` (13.11.2026, v3.0) ve 53 ekran görüntüsü.
Yöntem: çalışma kitabı açılıp içindeki 141 VBA bileşeni (≈52.700 satır, 1.389 prosedür) çıkarıldı ve statik olarak denetlendi; ekranlar tek tek incelendi. Excel ve SAP bu ortamda çalıştırılamadığı için çalışma zamanı davranışı ölçülmedi.

---

## 1. Sistem ne yapıyor?

AKSORA, Çoruh EDAŞ'ın beş ilindeki (Artvin, Rize, Trabzon, Gümüşhane, Giresun) malzeme stoğunu SAP'den okuyup **karar listesine** çeviren bir Excel uygulamasıdır. SAP "ne var"ı söyler; AKSORA "bugün ne yapılmalı"yı söyler.

| Birim | Sorduğu soru | Ekranlar |
|---|---|---|
| Malzeme Stok | Ne almalıyım, neyi almamalıyım, neyi başka ile göndermeliyim? | Satın Alma, Transfer Listesi, Atıl Stok, Stok Karşılaştırma, Malzeme Arama |
| Yatırım | Yüklenicideki malzeme teminatı aşıyor mu, proje keşfine göre ne eksik / fazla? | Teminat Takibi, Hakediş / Proje Malzeme Takibi |
| Müşteri | Seri numaralı malzeme (sayaç vb.) nerede? | Seri no arama, ilçe / yüklenici stokları |
| İlçe Yönetim | Sahadan sökülen malzemenin demontaj girişi yapıldı mı, İSG malzemesi zimmetlendi mi? | Aylık Demontaj, İSG Malzemeleri, Birleşik Stok |
| Yönetim Merkezi | Bugün en önemli 5 iş ne? | Şifreli tek ekran: 5 il × 4 birim |

Veri akışı: SAP GUI betikleri (MB52, MB51, ME5A/ME2N, IQ09) → gizli kayıt sayfaları (il × birim) → analiz (atıl stok, kritik stok, transfer eşleştirme) → ekranlar ve Excel/PDF çıktıları. İzleme isteyenler için SAP'ye bağlanmayan "YAYIN" kopyası üretiliyor.

Ekran görüntülerindeki ölçek: tekil stok değeri **1,39 milyar TL**; atıl stok 31,9 mn TL; bekleyen transfer kazancı 7,6 mn TL; iptal edilebilir talep/sipariş 1,75 mn TL; fazla teslim alınmış sipariş 2,07 mn TL.

---

## 2. Beş gözle değerlendirme

"Alıcı" ve "satıcı"yı, bu sistemi **ürün olarak satın alacak** ve **satacak** kişi olarak yorumladım.

### Alıcı (bu sistemi satın almayı düşünen kurum)
**Beğendiği:** Değer önerisi rakamla kanıtlanıyor: ekranda "şu 286 nakli yaparsan 7,6 mn TL satın alma yapmazsın" yazıyor. SAP'ye ek lisans / entegrasyon projesi gerektirmiyor.
**Soracağı sorular ve bugünkü cevap:**
- *"Geliştiren kişi ayrılırsa ne olur?"* — Tek kişi, tek dosya. Kod kaynak olarak sürümlenmiyor, belge yok. **En büyük satın alma engeli budur.**
- *"Kim neyi ne zaman yaptı?"* — Yalnız son birkaç yönetici girişi tutuluyordu (250 karakter). Denetim izi yoktu → **v3.1'de eklendi.**
- *"Geçen aya göre iyileştik mi?"* — Cevap yoktu; her rapor anlık fotoğraftı → **v3.1'de eklendi.**
- *"Aynı anda kaç kişi kullanır?"* — Bir kişi günceller, diğerleri YAYIN kopyasını izler. Gerçek çok kullanıcılı değil.

### Satıcı (bu sistemi pazarlayacak kişi)
**Elindeki koz:** "Bugün Yapılacak 5 İş" ekranı ve TL cinsinden tasarruf. Demo etkisi yüksek; görsel kimlik güçlü.
**Eksikleri:**
- Kanıtlanmış sonuç yok: önerilen tasarrufun ne kadarının **gerçekleştiği** ölçülmüyor. Satışta "fırsat" değil "gerçekleşen kazanç" satılır. Karne geçmişi (v3.1) bunun ilk adımı.
- Çoruh EDAŞ'a özel: il adları, üretim yeri kodları, depo kuralları kodun içinde. Başka bir dağıtım şirketine satmak için bunların bir ayar sayfasına taşınması gerekir.
- Yöneticinin toplantıya götüreceği tek sayfalık çıktı yoktu → **v3.1'de eklendi (PDF).**

### Kullanacak mühendis
**Beğendiği:** Tek tuşla güncelleme, anlaşılır Türkçe eylem dili ("KENDİ İHTİYACINDA KULLAN", "TRANSFER ET"), her listede Excel'e aktarma, hata olunca Excel'i toparlayan `CORE_EXCEL_KURTAR`.
**Zorlandığı:**
- Bir işlem başarısız olduğunda ne olduğunu sonradan göremiyor; hata mesajı kapatılınca kayboluyor → **v3.1: işlem günlüğü.**
- "Dosya mı bozuk, SAP mi kapalı, bilgisayarımda bir bileşen mi eksik?" sorusunun cevabı yoktu → **v3.1: sağlık kontrolü.**
- Dosya 24 MB, 166 sayfa; açılış ve kaydetme süresi büyüdükçe artacak.

### Holding sahibi
Baktığı tek şey: **bağlı sermaye ve risk.** 1,39 milyar TL stokun 31,9 mn TL'si 24 aydır hareketsiz; 179 kalem 1 aydan az yetiyor. Bu iki sayı aynı ekranda ve doğru.
**Eksik olan:** yön. "Atıl stok geçen çeyreğe göre azaldı mı, hangi il müdürü iyileştirdi?" sorusuna sistem cevap veremiyordu. v3.1 karnesi sekiz göstergeyi tarihe bağlıyor; birkaç hafta sonra bu soru cevaplanabilir olacak.
**Görmek isteyeceği ama henüz olmayan:** gerçekleşen tasarruf (TL), il müdürü bazında hedef / gerçekleşme.

### Müdür
**Beğendiği:** Demontaj girişi yapılmayan ilçe, zimmetlenmemiş İSG malzemesi, teminatı aşan yüklenici — hepsi isim isim. Takip edilebilir.
**Eksik olan:** Verilen işin **kapanıp kapanmadığını** sistem bilmiyor; aynı uyarı ertesi gün yine çıkıyor ama "3 gündür bekliyor" demiyor. Sorumlu ataması ve bekleme süresi yok (yol haritası, madde 3).

---

## 3. Teknik bulgular

| # | Bulgu | Önem | Durum |
|---|---|---|---|
| 1 | **Şifreler kodda düz metin** (`ADM_SIFRE` ve dört birim şifresi) ve depoya yüklenen VBA zip dosyalarında da duruyor. Depo herkese açıksa şifreler açıktadır. | Yüksek | Kaynak aynasında maskelendi. **Şifrelerin değiştirilmesi ve zip dosyalarının depodan çıkarılması sizin kararınız** (madde 5). |
| 2 | VBA projesi kilitsizse şifre kontrolü tek satırla atlanabilir; 104 yerde sayfa koruması parolasız (`Password:=""`). Koruma "yanlışlıkla bozmayı" önler, yetkilendirme değildir. | Yüksek | Sağlık kontrolü VBA kilidini raporluyor. Gerçek yetkilendirme dosya paylaşım izinleriyle yapılmalı. |
| 3 | Denetim izi yoktu. | Yüksek | **v3.1'de giderildi.** |
| 4 | Geçmiş / trend yoktu. | Yüksek | **v3.1'de giderildi.** |
| 5 | Kod sürümlenmiyordu: depoda yalnız `.xlsm` ve zip vardı; iki sürüm arasında ne değiştiği görülemiyordu. | Orta | **`src/vba/` kaynak aynası ve `tools/aksora_vba.py` eklendi.** |
| 6 | `modAKSORA_BIRIM` 11.278 satır; dört birimi, yönetim merkezini ve bekleme kartını birlikte içeriyor. `ADM_KUR` VBA'nın prosedür boyutu sınırına dayanmış (kodda not var). | Orta | Yol haritası: birim başına modüle bölme. |
| 7 | 739 `On Error Resume Next`. Çoğu görsel işlemde makul; hesaplama içinde olanlar yanlış sayıyı sessizce üretebilir. | Orta | Yeni kod hataları günlüğe yazıyor. Mevcut kod için ayrı bir tarama gerekir. |
| 8 | `VBScript.RegExp` kullanılıyor (6 yer). Microsoft VBScript'i Windows'tan aşamalı kaldırıyor. | Orta | Sağlık kontrolü izliyor. Yerleşik VBA RegExp'e geçilmeli. |
| 9 | 46 Windows API bildirimi (şifre kutusu maskeleme, pencere yönetimi). Office / Windows güncellemelerinde kırılmaya en yatkın bölüm. | Düşük | İzlenmeli. |
| 10 | Dosya 24 MB; metin tablosu açıkken 25 MB, sayfalar 61 MB. Her il × birim × gün kaydı sayfada tutuluyor; büyüme sürekli. | Orta | Sağlık kontrolü boyutu ve en büyük sayfayı raporluyor. Kalıcı çözüm: geçmişi ayrı dosyaya / veritabanına taşımak. |
| 11 | Otomatik test yok; 52 bin satır elle deneniyor. | Orta | v3.1 saf hesap işlevleri için sınama içeriyor; kurulum betiği bunu çalıştırmadan kaydetmiyor. |

Olumlu tespitler: tüm modüllerde `Option Explicit` var; `.Select/.Activate` kullanımı az (160); veri okuma dizilerle yapılıyor; uzun işlemlerde Excel durumu tek noktadan (`CORE_ISLEM_BASLAT/BITIR`) yönetiliyor; açılışta her adım ayrı korunuyor. Kod kalitesi, tek kişinin yazdığı bir Excel uygulaması için ortalamanın belirgin üstünde.

---

## 4. v3.1 ile eklenenler

Var olan hiçbir ekran ya da hesap değiştirilmedi. Üç yeni modül ve mevcut kodda **on tek satırlık bağlantı** var (`gelistirme/baglantilar.json`).

### Yönetim karnesi — trend (`modAKSORA_KPI`)
Yönetim Merkezi'nde "Bugün Yapılacak 5 İş"in altına yeni tablo. Her rapor yenilemesinde sekiz gösterge günde bir satır olarak kaydedilir (`_AKSORA_KPI`, çok gizli sayfa) ve önceki kayıtla / 30 gün öncesiyle karşılaştırılır:

sağlık skoru · tekil stok değeri · atıl stok değeri · atıl oranı · kritik stok kalemi · bekleyen tasarruf fırsatı · demontaj girişi eksiği · zimmetsiz İSG malzemesi

Her satırda değişim (▲/▼, TL ve %), eğilim çizgisi (son 12 kayıt) ve değerlendirme (İYİLEŞTİ / KÖTÜLEŞTİ). İl seçildiğinde o ilin kendi geçmişi tutulur. İlk gün "İLK KAYIT" yazar; karşılaştırma ertesi günden itibaren dolar.

### Yönetici özeti — tek sayfa PDF
Yönetim Merkezi üstündeki **YÖNETİCİ ÖZETİ (PDF)** düğmesi: göstergeler, bugün yapılacak işler, il karşılaştırması (stok payı, atıl oranı) ve öncelikli uyarılar tek A4 sayfada masaüstüne kaydedilir ve açılır.

### İşlem günlüğü (`modAKSORA_GUNLUK`)
Her SAP güncellemesi (süresiyle), kullanıcıya gösterilen her hata, her yönetici girişi (başarılı / hatalı) ve tek tuş güncelleme; kullanıcı ve bilgisayar adıyla kaydedilir (`_AKSORA_GUNLUK`, en fazla 5.000 satır). **İŞLEM GÜNLÜĞÜ** düğmesi Excel'e aktarır.

### Sistem sağlık kontrolü
**SAĞLIK KONTROLÜ** düğmesi: Excel sürümü, dosya boyutu ve konumu, zorunlu sayfalar, bozuk adlar, Windows bileşenleri, SAP GUI erişimi, VBA proje kilidi, son 7 günün hataları. Yalnız okur; sonucu TAMAM / UYARI / HATA olarak Excel'e aktarır.

### Akıcı bekleme kartı (`modAKSORA_AKICI`)
Kart yalnız VBA "ilerle" dediği anda çiziliyordu; SAP sorgusu sürerken (10-20 sn) çubuk ve geçen süre donuyordu. Kart açıkken çalışan bir Windows zamanlayıcısı artık çubuğun üzerinde bir şerit gezdiriyor ve "Geçen süre" yazısını saniye saniye yeniliyor. Zamanlayıcı yalnız kartı çizer; SAP'ye ve sayfalara dokunmaz. SAP çağrısı beklenirken Windows'un zamanlayıcıyı çalıştırıp çalıştırmadığı Microsoft belgelerinde net değil; bu davranış sizin bilgisayarınızda denenerek görülecek.

### Depo düzeni
- `src/vba/` — kodun okunabilir aynası (şifreler maskeli). Sürümler arası fark artık görülebilir.
- `tools/aksora_vba.py` — kodu çıkarma, bağlantıları doğrulama, statik denetim.
- `gelistirme/` — kurulum paketi.

### Doğrulama durumu (açık söylemek gerekirse)
- On bağlantının çapa satırı gerçek kodda tam olarak bir kez bulundu.
- Yeni modüller mevcut kodla ad çakışması, tanımsız değişken ve parametre sayısı açısından tarandı; saf hesap işlevleri LibreOffice Basic'te (VBA kipi) çalıştırılıp sınamadan geçti.
- **Excel'de derlenmedi ve çalıştırılmadı; kurulum betiği (PowerShell) de bu ortamda çalıştırılamadı.** Bu yüzden kurulum önce yedek alır, kodu ekler, derleme sınamasını çalıştırır ve yalnız "OK" gelirse kaydeder. İlk kurulumu dosyanın bir kopyasında yapın.

---

## 5. Sizin kararınızı bekleyenler

1. **Şifreler:** beş şifreyi değiştirin; `AKSORA_VBA_*.zip` ve `GÜNCEL VBA.zip` dosyalarını depodan çıkarın (geçmişte kalırlar; depo herkese açıksa geçmişin de temizlenmesi gerekir). Şirket verisi içeren `.xlsm`'in depoda durması da ayrıca değerlendirilmeli.
2. **VBA projesine parola verin** (Araçlar > VBAProject Özellikleri > Koruma).

## 6. Yol haritası (öncelik sırasıyla)

1. **Gerçekleşen tasarruf:** önerilen transfer / iptal sonraki okumada gerçekleşmişse "kazanıldı" olarak say. Holding ve satış için en değerli tek sayı.
2. **Ayar sayfası:** il, üretim yeri, depo kuralları ve eşikleri koddan çıkar. Başka şirkete uyarlanabilirlik buna bağlı.
3. **İş takibi:** her uyarıya "ilk görülme tarihi" ve bekleme günü; müdür için "3 gündür kapanmadı" listesi.
4. **Geçmişi dosyadan ayır:** günlük stok kayıtlarını ayrı bir veri dosyasına taşı; ana dosya küçük kalsın.
5. **`modAKSORA_BIRIM`'i böl:** birim başına modül; `ADM_KUR` çizim bölümleri ayrı işlevlere.
6. **Web paneli:** YAYIN kopyası yerine salt okunur web ekranı (telefon / tablet).
