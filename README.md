# AKSORA

Çoruh EDAŞ — beş il (Artvin, Rize, Trabzon, Gümüşhane, Giresun) için SAP tabanlı stok, satın alma, yatırım ve ilçe yönetimi karar destek uygulaması (Excel / VBA).

| Klasör | İçerik |
|---|---|
| `AKSORA.xlsm` | Uygulama |
| `docs/INCELEME_RAPORU.md` | Sistem incelemesi, bulgular, v3.1 geliştirmeleri, yol haritası |
| `gelistirme/` | v3.1 geliştirme paketi (yeni modüller + kurulum) |
| `src/vba/` | VBA kodunun okunabilir aynası (v3.1 hali; şifreler maskeli, yalnız inceleme için) |
| `tools/aksora_vba.py` | Kod çıkarma, bağlantı doğrulama, statik denetim |

## v3.1 kurulumu

Paket Excel'de henüz derlenip denenmedi; **ilk kurulumu `AKSORA.xlsm`'in bir kopyasında yapın.**

1. Excel'de bir kez: Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarları > Makro Ayarları > **"VBA proje nesne modeline erişime güven"** işaretli olsun.
2. AKSORA kapalıyken `gelistirme\KUR.cmd` dosyasına çift tıklayın ve `.xlsm` dosyasını seçin.
3. Betik yedek alır, modülleri ve on bağlantı satırını ekler, derleme sınamasını çalıştırır; yalnız sınama `OK` dönerse kaydeder. Aksi halde dosyaya dokunmadan kapanır.

Kurulumdan sonra: ADMİN GİRİŞİ > Yönetim Merkezi'nde **Yönetim Karnesi — Trend** tablosu ve üstte **YÖNETİCİ PANOSU (WEB)**, **YÖNETİCİ ÖZETİ (PDF)**, **İŞLEM GÜNLÜĞÜ**, **SAĞLIK KONTROLÜ** düğmeleri görünür. Bekleme kartında (ör. TÜM BİRİMLERİ GÜNCELLE) çubuğun üzerinde kayan bir şerit ve saniye saniye ilerleyen süre görünür.

Betik çalışmazsa (kurumsal kısıt): `KUR.cmd -Elle` komutu `gelistirme\elle_kurulum\` klasörüne içe aktarılacak modülleri ve elle eklenecek satırların listesini yazar.

Geri alma: betiğin aynı klasöre bıraktığı `*_yedek_*.xlsm` dosyasını geri kopyalayın.

## Araçlar

```bash
pip install oletools
```

```bash
python3 tools/aksora_vba.py cikar AKSORA.xlsm src/vba --maskele
```

```bash
python3 tools/aksora_vba.py bagla src/vba gelistirme/baglantilar.json
```

```bash
python3 tools/aksora_vba.py denetle src/vba
```
