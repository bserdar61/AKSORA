param(
    [string]$Dosya = "",
    [switch]$Elle
)

# ==========================================================
# AKSORA v3.1 gelistirme paketi kurulumu
#
# Ne yapar:
#   1. Sectiginiz AKSORA.xlsm dosyasinin yedegini alir (ayni klasore).
#   2. Dosyayi makrolari CALISTIRMADAN acar (olaylar kapali).
#   3. moduller\ klasorundeki yeni modulleri ekler (varsa eskisini degistirir).
#   4. baglantilar.json icindeki tek satirlik baglantilari ilgili modullere ekler.
#      Her baglanti icin capa satiri tam olarak bir kez bulunmalidir; aksi halde kurulum durur.
#   5. AKSORA_GELISTIRME_SINAMA makrosunu calistirir (yeni kod derleniyor mu?).
#   6. Yalniz sinama "OK" donerse kaydeder. Aksi halde KAYDETMEDEN kapatir.
#
# Gereken: Excel > Dosya > Secenekler > Guven Merkezi > Guven Merkezi Ayarlari >
#          Makro Ayarlari > "VBA proje nesne modeline erisime guven" isaretli olmali.
#
# -Elle : Excel'e dokunmaz; elle kurulum icin ANSI (Windows-1254) modul dosyalarini
#         ve eklenecek satirlarin listesini "elle_kurulum" klasorune yazar.
# ==========================================================

$ErrorActionPreference = "Stop"
$kok = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulKlasor = Join-Path $kok "moduller"
$jsonYol = Join-Path $kok "baglantilar.json"
$ansi = [System.Text.Encoding]::GetEncoding(1254)
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Yaz([string]$metin, [string]$renk = "Gray") {
    Write-Host $metin -ForegroundColor $renk
}

function AnsiKopya([string]$kaynak, [string]$hedef) {
    # Depodaki moduller UTF-8'dir; VBA duzenleyicisi Windows-1254 ve CRLF bekler.
    $metin = [System.IO.File]::ReadAllText($kaynak, $utf8)
    $metin = $metin -replace "`r`n", "`n"
    $metin = $metin -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText($hedef, $metin, $ansi)
}

function BaglantiYeri([string[]]$satirlar, $b) {
    # Donus: 0 = zaten var; pozitif = eklenecek satir numarasi (1 tabanli); hata durumunda throw.
    $hedef = $b.satir.Trim().ToLowerInvariant()
    foreach ($s in $satirlar) {
        if ($s.Trim().ToLowerInvariant() -eq $hedef) { return 0 }
    }
    $capa = $b.capa.Trim().ToLowerInvariant()
    $bulunan = @()
    for ($i = 0; $i -lt $satirlar.Length; $i++) {
        $t = $satirlar[$i].Trim().ToLowerInvariant()
        if ($b.eslesme -eq "tam") {
            if ($t -eq $capa) { $bulunan += $i }
        } else {
            if ($t.StartsWith($capa)) { $bulunan += $i }
        }
    }
    if ($bulunan.Count -ne 1) {
        throw ("Capa satiri " + $bulunan.Count + " kez bulundu (1 olmali): " + $b.capa)
    }
    $i = $bulunan[0]
    if ($satirlar[$i].TrimEnd().EndsWith(" _")) { throw ("Capa satiri devam satiri: " + $b.capa) }
    if ($b.konum -eq "once" -and $i -gt 0 -and $satirlar[$i - 1].TrimEnd().EndsWith(" _")) {
        throw ("Capa bir devam satirinin parcasi: " + $b.capa)
    }
    if ($b.konum -eq "sonra") { return ($i + 2) } else { return ($i + 1) }
}

Yaz ""
Yaz "AKSORA v3.1 gelistirme paketi kurulumu" "Cyan"
Yaz "--------------------------------------" "Cyan"

if (-not (Test-Path $modulKlasor)) { throw "moduller klasoru bulunamadi: $modulKlasor" }
if (-not (Test-Path $jsonYol)) { throw "baglantilar.json bulunamadi: $jsonYol" }
$baglantilar = [System.IO.File]::ReadAllText($jsonYol, $utf8) | ConvertFrom-Json
$moduller = Get-ChildItem -Path $modulKlasor -Filter "*.bas" | Sort-Object Name

# ---------- elle kurulum dosyalari ----------
if ($Elle) {
    $elleKlasor = Join-Path $kok "elle_kurulum"
    if (-not (Test-Path $elleKlasor)) { New-Item -ItemType Directory -Path $elleKlasor | Out-Null }
    foreach ($m in $moduller) { AnsiKopya $m.FullName (Join-Path $elleKlasor $m.Name) }
    $liste = New-Object System.Collections.Generic.List[string]
    $liste.Add("AKSORA v3.1 - elle eklenecek satirlar")
    $liste.Add("1) VBA duzenleyicisi (Alt+F11) > Dosya > Dosya Al ile bu klasordeki .bas dosyalarini alin.")
    $liste.Add("2) Asagidaki her satiri, belirtilen moduldeki capa satirinin oncesine / sonrasina ekleyin.")
    $liste.Add("")
    foreach ($b in $baglantilar) {
        $liste.Add("Modul   : " + $b.modul)
        $liste.Add("Capa    : " + $b.capa + $(if ($b.eslesme -eq "onek") { "  (bu metinle baslayan satir)" } else { "" }))
        $liste.Add("Konum   : capa satirindan " + $(if ($b.konum -eq "sonra") { "SONRA" } else { "ONCE" }))
        $liste.Add("Eklenecek: " + $b.satir)
        $liste.Add("")
    }
    [System.IO.File]::WriteAllLines((Join-Path $elleKlasor "EKLENECEK_SATIRLAR.txt"), $liste, $ansi)
    Yaz "Elle kurulum dosyalari yazildi: $elleKlasor" "Green"
    exit 0
}

# ---------- dosya secimi ----------
if ([string]::IsNullOrWhiteSpace($Dosya)) {
    Add-Type -AssemblyName System.Windows.Forms
    $pencere = New-Object System.Windows.Forms.OpenFileDialog
    $pencere.Title = "AKSORA calisma kitabini secin"
    $pencere.Filter = "Excel makro kitabi (*.xlsm)|*.xlsm"
    if ($pencere.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Yaz "Dosya secilmedi; kurulum iptal edildi." "Yellow"
        exit 1
    }
    $Dosya = $pencere.FileName
}
if (-not (Test-Path $Dosya)) { throw "Dosya bulunamadi: $Dosya" }
$Dosya = (Resolve-Path $Dosya).Path

# Dosya baska bir Excel'de acik mi?
try {
    $akis = [System.IO.File]::Open($Dosya, "Open", "ReadWrite", "None")
    $akis.Close()
} catch {
    throw "Dosya su anda acik ya da salt okunur. AKSORA'yi kapatip yeniden deneyin."
}

# ---------- yedek ----------
$yedek = Join-Path (Split-Path -Parent $Dosya) (
    [System.IO.Path]::GetFileNameWithoutExtension($Dosya) + "_yedek_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsm")
Copy-Item -Path $Dosya -Destination $yedek
Yaz "Yedek alindi : $yedek" "Green"

# ---------- gecici ANSI moduller ----------
$gecici = Join-Path ([System.IO.Path]::GetTempPath()) ("aksora_kur_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $gecici | Out-Null
foreach ($m in $moduller) { AnsiKopya $m.FullName (Join-Path $gecici $m.Name) }

$xl = $null
$wb = $null
$kaydedildi = $false
try {
    Yaz "Excel aciliyor (makrolar calistirilmaz)..."
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $xl.AskToUpdateLinks = $false
    $wb = $xl.Workbooks.Open($Dosya, 0, $false)
    if ($wb.ReadOnly) { throw "Dosya salt okunur acildi; kurulum yapilamaz." }

    try {
        $proje = $wb.VBProject
        $bilesenler = $proje.VBComponents
        $adet = $bilesenler.Count
    } catch {
        throw ("VBA projesine erisilemedi. Excel > Dosya > Secenekler > Guven Merkezi > Guven Merkezi Ayarlari > " +
            "Makro Ayarlari > 'VBA proje nesne modeline erisime guven' secenegini isaretleyip yeniden deneyin.")
    }
    if ($proje.Protection -ne 0) { throw "VBA projesi parola ile kilitli. Once kilidi acin (Alt+F11), sonra yeniden deneyin." }
    Yaz "VBA projesi   : $adet bilesen"

    # ---------- 1) once baglantilarin hepsi denenir (hicbir sey degistirilmeden) ----------
    $plan = @()
    foreach ($b in $baglantilar) {
        try { $bilesen = $bilesenler.Item($b.modul) } catch { throw ("Modul bulunamadi: " + $b.modul) }
        $kod = $bilesen.CodeModule
        $satirlar = $kod.Lines(1, $kod.CountOfLines) -split "`r`n"
        $yer = BaglantiYeri $satirlar $b
        $plan += [pscustomobject]@{ Baglanti = $b; Yer = $yer }
    }
    Yaz "Baglantilar   : $($plan.Count) capa dogrulandi"

    # ---------- 2) moduller ----------
    foreach ($m in $moduller) {
        $ad = [System.IO.Path]::GetFileNameWithoutExtension($m.Name)
        $eski = $null
        try { $eski = $bilesenler.Item($ad) } catch { $eski = $null }
        if ($eski -ne $null) {
            $eski.Name = $ad + "_ESKI"
            $bilesenler.Remove($eski)
            Yaz "  guncellendi : $ad"
        } else {
            Yaz "  eklendi     : $ad"
        }
        [void]$bilesenler.Import((Join-Path $gecici $m.Name))
    }

    # ---------- 3) baglantilar (her eklemeden once yer yeniden hesaplanir) ----------
    foreach ($p in $plan) {
        $b = $p.Baglanti
        $kod = $bilesenler.Item($b.modul).CodeModule
        $satirlar = $kod.Lines(1, $kod.CountOfLines) -split "`r`n"
        $yer = BaglantiYeri $satirlar $b
        if ($yer -eq 0) {
            Yaz ("  zaten var   : " + $b.modul + "  |  " + $b.satir.Trim())
        } else {
            $kod.InsertLines($yer, $b.satir)
            Yaz ("  baglandi    : " + $b.modul + " satir " + $yer + "  |  " + $b.satir.Trim())
        }
    }

    # ---------- 4) sinama ----------
    Yaz "Sinama calistiriliyor (derleme kontrolu)..."
    Yaz "  Excel'de 'Derleme hatasi' penceresi cikarsa pencereyi kapatin; dosya kaydedilmeyecektir." "DarkGray"
    $makro = "'" + $wb.Name.Replace("'", "''") + "'!AKSORA_GELISTIRME_SINAMA"
    $sonuc = [string]$xl.Run($makro)
    if ($sonuc -ne "OK") { throw ("Sinama basarisiz: " + $sonuc) }
    Yaz "Sinama        : OK" "Green"

    # ---------- 5) kaydet ----------
    $wb.Save()
    $kaydedildi = $true
    Yaz "Kaydedildi    : $Dosya" "Green"
}
catch {
    Yaz ""
    Yaz ("KURULUM YAPILAMADI: " + $_.Exception.Message) "Red"
    Yaz "Dosyaniz DEGISTIRILMEDI (kaydedilmeden kapatildi). Yedek de duruyor." "Yellow"
}
finally {
    try { if ($wb -ne $null) { $wb.Close($false) } } catch { }
    try { if ($xl -ne $null) { $xl.Quit() } } catch { }
    if ($wb -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($xl -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    [System.GC]::Collect()
    try { Remove-Item -Path $gecici -Recurse -Force } catch { }
}

if ($kaydedildi) {
    Yaz ""
    Yaz "KURULUM TAMAMLANDI." "Green"
    Yaz "AKSORA'yi acin > ADMIN GIRISI > Yonetim Merkezi:"
    Yaz "  - 'YONETIM KARNESI - TREND' tablosu (Bugun Yapilacak Isler'in altinda)"
    Yaz "  - Ustte: YONETICI OZETI (PDF), ISLEM GUNLUGU, SAGLIK KONTROLU dugmeleri"
    Yaz "Sorun olursa yedek dosyayi geri kopyalamaniz yeterlidir: $yedek"
    exit 0
}
exit 1
