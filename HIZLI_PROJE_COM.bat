@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title COM Automation - MS Project Proje Oluşturucu

echo 🚀 COM AUTOMATION - MS PROJECT PROJE OLUŞTURUCU
echo =========================================================
echo 📅 Proje: 28.07.2025 → 31.10.2025 (3 Ay)
echo 🏗️ Paralel çalışma: 5 alan eş zamanlı  
echo 🤖 Tam MS Project COM otomasyonu
echo ⚡ Excel + MPP dosyaları tek seferde oluşturulur
echo.

echo [1/5] 🔧 Sistem kontrolleri...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ❌ HATA: Python yüklü değil!
    echo 💡 Python'u yükleyin: https://python.org
    pause
    exit /b 1
)
echo    ✅ Python hazır

echo.
echo [2/5] 📦 COM automation paketleri kontrol ediliyor...
echo    📦 comtypes ve openpyxl paketleri kontrol ediliyor...
python -c "import comtypes.client; print('   ✅ comtypes hazır')" 2>nul || (
    echo    📥 comtypes paketi yükleniyor...
    pip install comtypes
    if %ERRORLEVEL% neq 0 (
        echo    ❌ comtypes yüklenemedi
        goto FALLBACK_MODE
    )
    echo    ✅ comtypes yüklendi
)

python -c "import openpyxl; print('   ✅ openpyxl hazır')" 2>nul || (
    echo    📥 openpyxl paketi yükleniyor...
    pip install openpyxl
    if %ERRORLEVEL% neq 0 (
        echo    ❌ openpyxl yüklenemedi
        pause
        exit /b 1
    )
    echo    ✅ openpyxl yüklendi
)

echo.
echo [3/5] 🤖 Gelişmiş COM template oluşturuluyor...
echo    🔄 Excel şablonu + MS Project dosyası birlikte oluşturuluyor...
python com_template_creator.py
if %ERRORLEVEL% EQU 0 (
    echo    ✅ COM template oluşturma başarılı!
    goto COM_SUCCESS
) else (
    echo    ⚠️ COM template oluşturma başarısız, gelişmiş automation deneniyor...
    goto ADVANCED_COM
)

:ADVANCED_COM
echo.
echo [4/5] 🚀 Gelişmiş COM automation başlatılıyor...
echo    🔧 MS Project COM API kullanarak proje oluşturuluyor...
python advanced_com_automation.py
if %ERRORLEVEL% EQU 0 (
    echo    ✅ Gelişmiş COM automation başarılı!
    goto COM_SUCCESS
) else (
    echo    ⚠️ Gelişmiş COM automation başarısız, fallback mode'a geçiliyor...
    goto FALLBACK_MODE
)

:FALLBACK_MODE
echo.
echo [4/5] 🔄 Fallback mode: Basit Excel şablonu oluşturuluyor...
python create_simple_template.py
if %ERRORLEVEL% neq 0 (
    echo    ❌ HATA: Excel şablonu oluşturulamadı!
    pause
    exit /b 1
)
echo    ✅ Excel şablonu hazır

echo.
echo [5/5] 🔄 Manuel MPP dönüştürme deneniyor...
python comtypes_excel_to_msp.py
if %ERRORLEVEL% EQU 0 (
    echo    ✅ Manuel MPP dönüştürme başarılı
    goto SUCCESS_MPP
) else (
    echo    ⚠️ MPP oluşturma başarısız - Excel şablonu kullanılabilir
    goto SUCCESS_EXCEL
)

:COM_SUCCESS
echo.
echo [5/5] 🎯 COM automation dosyaları kontrol ediliyor...
if exist "data\SporSalonu_Optimized_26_07_2025.mpp" (
    echo    ✅ MS Project dosyası oluşturuldu
    goto SUCCESS_COM_FULL
) else (
    echo    ✅ Excel şablonu oluşturuldu
    goto SUCCESS_EXCEL
)

:SUCCESS_COM_FULL
echo.
echo ================================================================
echo                     🎉 TAM COM AUTOMATION BAŞARILI! 🎉
echo ================================================================
echo  📅 Başlangıç: 28.07.2025 (Pazartesi)
echo  📅 Bitiş: 31.10.2025 (3 Ay)
echo  🏗️ 5 Alan Paralel Çalışma Strategy
echo  🤖 MS Project COM API Entegrasyonu
echo  👥 30 Optimize Görev + 14 Kaynak
echo.
echo  🎯 OLUŞTURULAN DOSYALAR:
echo  📊 Excel Şablonu: data\proje_sablonu.xlsx
echo  📈 MS Project MPP: data\SporSalonu_Optimized_26_07_2025.mpp
echo.
echo  ⚡ ÖZELLİKLER:
echo  • Otomatik görev bağımlılıkları
echo  • Kaynak atamaları ve maliyetleri  
echo  • Kritik yol analizi
echo  • Paralel çalışma optimizasyonu
echo  • Gantt Chart görselleştirme
echo  • Resource leveling
echo ================================================================
echo.

echo 📂 MS Project dosyasını açmak için Enter'a basın...
pause >nul
if exist "data\SporSalonu_Optimized_26_07_2025.mpp" (
    start "" "data\SporSalonu_Optimized_26_07_2025.mpp"
) else (
    echo ❌ MPP dosyası bulunamadı, Excel açılıyor...
    start "" "data\proje_sablonu.xlsx"
)
goto END

:SUCCESS_MPP
echo.
echo ================================================================
echo                  ✅ MS PROJECT AUTOMATION BAŞARILI! ✅
echo ================================================================
echo  📅 Başlangıç: 28.07.2025 (Pazartesi)
echo  📅 Bitiş: 31.10.2025 (3 Ay)
echo  🏗️ 5 Alan Paralel Çalışma
echo  👥 Optimize Görev Planlaması
echo.
echo  📁 OLUŞTURULAN DOSYALAR:
echo  📊 Excel: data\proje_sablonu.xlsx
echo  📈 MPP: data\SporSalonu_Optimized_26_07_2025.mpp
echo.
echo  🎯 MS Project dosyası hazır!
echo ================================================================
echo.
echo 📂 Proje dosyasını açmak için Enter'a basın...
pause >nul
start "" "data\SporSalonu_Optimized_26_07_2025.mpp"
goto END

:SUCCESS_EXCEL
echo.
echo ================================================================
echo                    ✅ EXCEL ŞABLONU HAZIR! ✅
echo ================================================================
echo  📅 Başlangıç: 28.07.2025 (Pazartesi)
echo  📅 Bitiş: 31.10.2025 (3 Ay)
echo  🏗️ 5 Alan Paralel Çalışma
echo  👥 Optimize Görev Planlaması
echo.
echo  📊 Excel şablonu başarıyla oluşturuldu!
echo  📁 Dosya: data\proje_sablonu.xlsx
echo.
echo  📝 MS PROJECT'E MANUEL AKTARIM:
echo  1. Microsoft Project'i açın
echo  2. Dosya ^> Aç ^> Excel dosyasını seçin
echo  3. İçe aktarma sihirbazını takip edin
echo  4. Görev Adı, Süre, Başlangıç eşleştirin
echo  5. MPP olarak kaydedin
echo.
echo  💡 İPUCU: COM automation için Microsoft
echo     Project'in yüklü ve lisanslı olması gerekir
echo ================================================================
echo.
echo 📂 Excel dosyasını açmak için Enter'a basın...
pause >nul
start "" "data\proje_sablonu.xlsx"

:END
echo.
echo 📚 OLUŞTURULAN DOSYALAR:
echo ================================
if exist "data\proje_sablonu.xlsx" (
    echo ✅ Excel Şablonu: data\proje_sablonu.xlsx
)
if exist "data\SporSalonu_Optimized_26_07_2025.mpp" (
    echo ✅ MS Project MPP: data\SporSalonu_Optimized_26_07_2025.mpp
)
if exist "data\SporSalonu_Optimized_26_07_2025.csv" (
    echo ✅ CSV Export: data\SporSalonu_Optimized_26_07_2025.csv
)
if exist "data\SporSalonu_Optimized_26_07_2025.xml" (
    echo ✅ XML Export: data\SporSalonu_Optimized_26_07_2025.xml
)
echo ✅ Kullanım Kılavuzu: KULLANIM_KILAVUZU.md
echo ================================
echo.
echo 🎓 KULLANIM SENARYOLARI:
echo • COM Automation: Tam otomatik MPP oluşturma
echo • Excel Manual: Excel'den MS Project'e manuel aktarım
echo • XML/CSV Export: Alternatif format desteği
echo.
echo 🔧 Sorun mu yaşıyorsunuz?
echo • Python'un PATH'e ekli olduğundan emin olun
echo • Microsoft Project'in yüklü olduğunu kontrol edin
echo • Detaylar için KULLANIM_KILAVUZU.md dosyasını inceleyin
echo.
echo ✅ İşlem tamamlandı! COM automation sistemi aktif.

pause
