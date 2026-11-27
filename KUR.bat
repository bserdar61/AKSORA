@echo off
setlocal EnableExtensions
title AKSORA Kurulum

rem ============================================================
rem  AKSORA kurulum: masaustu ve baslat menusu simgesi olusturur.
rem  Bu dosyayi AKSORA klasorunun icine koyup cift tiklayin.
rem  Hicbir dosya silinmez, yonetici izni gerekmez.
rem ============================================================

rem --- Bu .bat dosyasinin bulundugu klasor = AKSORA klasoru ---
set "AKS_KLASOR=%~dp0"
if "%AKS_KLASOR:~-1%"=="\" set "AKS_KLASOR=%AKS_KLASOR:~0,-1%"
set "AKS_HEDEF=%AKS_KLASOR%\AKSORA_BASLAT.hta"
set "AKS_IKON=%AKS_KLASOR%\GORSELLER\LOGO\AKSORA.ico"
if not exist "%AKS_IKON%" set "AKS_IKON=%AKS_KLASOR%\AKSORA.ico"

if not exist "%AKS_HEDEF%" goto yokHata
if not exist "%AKS_KLASOR%\AKSORA*.xlsm" goto yokXlsm

echo.
echo   AKSORA simgeleri olusturuluyor...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $t=$env:AKS_HEDEF; $wd=$env:AKS_KLASOR; $ic=$env:AKS_IKON; foreach($d in @([Environment]::GetFolderPath('Desktop'),[Environment]::GetFolderPath('Programs'))){ $p=Join-Path $d 'AKSORA.lnk'; $s=$w.CreateShortcut($p); $s.TargetPath=$t; $s.WorkingDirectory=$wd; $s.Description='AKSORA'; if(Test-Path $ic){ $s.IconLocation=$ic+',0' }; $s.Save() }"
if errorlevel 1 goto psHata

echo   Tamamlandi.
powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show('AKSORA kuruldu.' + [char]10 + 'Masaustundeki AKSORA simgesinden acabilirsiniz.','AKSORA Kurulum','OK','Information')"
exit /b 0

:yokHata
echo.
echo   HATA: AKSORA_BASLAT.hta bulunamadi.
echo   KUR.bat dosyasini AKSORA klasorunun (AKSORA.xlsm ile ayni yer) icine
echo   koyup oradan calistirin.
echo.
pause
exit /b 1

:yokXlsm
echo.
echo   HATA: Bu klasorde AKSORA.xlsm bulunamadi.
echo   KUR.bat'i AKSORA klasorunun icine koyup calistirin.
echo.
pause
exit /b 1

:psHata
echo.
echo   Simgeler olusturulamadi (PowerShell engellenmis olabilir).
echo   Masaustu simgesi, AKSORA'yi ilk kez actiginizda otomatik olusur.
echo.
pause
exit /b 1
