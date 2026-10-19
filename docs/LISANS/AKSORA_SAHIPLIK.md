# AKSORA — Sahiplik ve Bütünlük Kaydı

**Eser:** AKSORA — Stok ve Verimlilik Uygulaması (Excel/VBA)
**Sahip / Geliştirici:** Burhan SERDAR
**Kurum:** Çoruh EDAŞ
**Bu kaydın tarihi:** 2026-10-19 08:46 UTC

Bu dosya, AKSORA'nın bu sürümünün Burhan SERDAR tarafından geliştirildiğini
belgeler. Uygulamanın tüm geliştirme geçmişi, tarih ve saatleriyle birlikte
bu deponun (git) kayıtlarında tutulmaktadır; her sürüm ayrıca aşağıdaki
SHA-256 parmak izleriyle kimliklenir.

## Dosya parmak izleri (SHA-256)

| Dosya | SHA-256 |
|---|---|
| AKSORA.xlsm (ana kopya) | `9aac5c97ccbc32d2a736a74d0996b6b90c0dad6cc3a4499115dd118b9ca36474` |
| AKSORA_DAGITIM.xlsm (paylaşım kopyası) | `9496cb3cc1f057007017a0c86f2da380c29bcbafe4b7da0ce72c153a71be1322` |

## Doğrulama

Herhangi bir kopyanın bu sürümle birebir aynı olup olmadığı,
Windows PowerShell'de şu komutla kontrol edilebilir:

```
Get-FileHash .\AKSORA.xlsm -Algorithm SHA256
```

Çıkan değer yukarıdaki parmak iziyle aynıysa dosya değiştirilmemiştir.

## Gömülü imzalar

- Açılış ekranı ve panel alt bilgilerinde görünür imza.
- Kod içinde `AKSORA_SAHIBI = "Burhan SERDAR"` sabiti.
- Gizli tanımlı adlar: `_AKSORA_SAHIP`, `_AKSORA_IMZA`.
- Kitap özellikleri (Yazar / Şirket).
- Her açılışta `AKSORA_IMZA_DOGRULA` imzaları koddan yeniden yazar;
  silinir veya değiştirilirse sahip adına geri döner.
