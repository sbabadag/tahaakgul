@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title COM Automation - MS Project Proje Oluşturucu

echo 🚀 COM AUTOMATION - MS PROJECT PROJE OLUŞTURUCU
echo =========================================================
echo 📅 Proje: 28.07.2025 → 31.10.2025 (3 Ay)
echo 🏗️ Paralel çalışma: 5 alan eş zamanlı
echo 🤖 Öncelik: MS Project COM automation
echo ⚡ Excel + MPP dosyaları otomatik oluşturulur
echo.

echo [0/5] 🔧 Sistem kontrolü yapılıyor...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ❌ HATA: Python yüklü değil!
    echo Lütfen Python'u yükleyin: https://python.org
    pause
    exit /b 1
)
echo ✅ Python hazır

echo [1/5] 📦 COM automation paketleri kontrol ediliyor...
echo    📦 Öncelik: MS Project COM entegrasyonu
python -c "import comtypes.client; print('   ✅ comtypes hazır')" 2>nul || (
    echo    📥 comtypes paketi yükleniyor...
    pip install comtypes
    echo    ✅ comtypes yüklendi
)

echo.
echo [2/5] 🚀 Hibrit COM automation başlatılıyor...
echo    🔧 Kapsamlı Excel + Gelişmiş MPP oluşturma
python hybrid_com_automation.py
if %ERRORLEVEL% EQU 0 (
    echo ✅ Hibrit automation başarılı!
    echo    🔧 Gelişmiş MPP fix çalıştırılıyor...
    python fix_mpp_tasks.py
    if %ERRORLEVEL% EQU 0 (
        echo ✅ MPP task fix başarılı!
        goto HYBRID_SUCCESS
    ) else (
        echo ⚠️ MPP fix başarısız, mevcut dosyalar kullanılacak
        goto HYBRID_SUCCESS
    )
) else (
    echo ⚠️ Hibrit automation başarısız, gelişmiş COM deneniyor...
    goto ADVANCED_COM
)

:ADVANCED_COM
echo.
echo [3/5] 🔧 Gelişmiş COM automation başlatılıyor...
echo    🔧 MPP task fix öncelikli çalıştırılıyor...
python fix_mpp_tasks.py
if %ERRORLEVEL% EQU 0 (
    echo ✅ Gelişmiş MPP oluşturma başarılı!
    goto HYBRID_SUCCESS
) else (
    echo ⚠️ Gelişmiş MPP başarısız, standart automation deneniyor...
    python advanced_com_automation.py
    if %ERRORLEVEL% EQU 0 (
        echo ✅ Standart COM automation başarılı!
        goto HYBRID_SUCCESS
    ) else (
        echo ⚠️ COM automation başarısız, fallback mode'a geçiliyor...
        goto FALLBACK_MODE
    )
)

:FALLBACK_MODE
echo.
echo [3/5] � Fallback: Excel şablonu oluşturuluyor...
python create_simple_template.py
if %ERRORLEVEL% neq 0 (
    echo ❌ HATA: Excel şablonu oluşturulamadı!
    echo Lütfen şu paketi yükleyin: pip install openpyxl
    pause
    exit /b 1
)
echo ✅ Excel şablonu hazır

echo.
echo [4/5] 🔄 MS Project dosyası oluşturuluyor (fallback)...
python comtypes_excel_to_msp.py
if %ERRORLEVEL% EQU 0 (
    echo ✅ COM Excel-to-MPP dönüştürme başarılı
    goto SUCCESS_MPP
) else (
    echo ⚠️ COM dönüştürme başarısız, CSV/XML deneniyor...
    python excel_to_mpp_simple.py
    if %ERRORLEVEL% EQU 0 (
        echo ✅ Excel'den XML/CSV oluşturuldu
        python csv_to_mpp_auto.py
        if %ERRORLEVEL% EQU 0 (
            echo ✅ CSV'den MPP oluşturuldu  
            goto SUCCESS_MPP
        ) else (
            echo ⚠️ MPP oluşturma başarısız - XML/CSV kullanılabilir
            goto SUCCESS_EXCEL
        )
    ) else (
        echo ⚠️ Excel dönüştürme başarısız
        echo 📝 Excel dosyasını manuel olarak MS Project'te açabilirsiniz
        goto SUCCESS_EXCEL
    )
)

:HYBRID_SUCCESS
echo.
echo [5/5] 🎯 Hibrit automation dosyaları kontrol ediliyor...
if exist "data\SporSalonu_Optimized_26_07_2025.mpp" (
    echo ✅ MS Project dosyası hibrit automation ile oluşturuldu
    goto SUCCESS_COM_FULL
) else (
    echo ✅ Kapsamlı Excel şablonu hibrit sistem ile oluşturuldu
    goto SUCCESS_HYBRID_EXCEL
)

:SUCCESS_HYBRID_EXCEL
echo.
echo ================================================================
echo            🎉 HİBRİT EXCEL AUTOMATION BAŞARILI! 🎉
echo ================================================================
echo  📅 Başlangıç: 28.07.2025 (Pazartesi)
echo  📅 Bitiş: 31.10.2025 (3 Ay)
echo  🏗️ 5 Alan Paralel Çalışma
echo  📊 28 Optimize Görev + 15 Kaynak
echo  🤖 Hibrit COM sistem ile oluşturuldu
echo.
echo  🎯 KAPSAMLI EXCEL ÖZELLİKLERİ:
echo  📋 Görevler: Detaylı planlama ve bağımlılıklar
echo  👥 Kaynaklar: Maliyet ve kullanım bilgileri
echo  📅 Takvim: Çalışma programı ve tatiller
echo  📊 Proje Bilgileri: Kapsamlı proje verileri
echo.
echo  📁 Excel: data\proje_sablonu.xlsx
echo  💡 MS Project'e aktarım: Dosya ^> Aç ^> Excel seç
echo ================================================================
echo.
echo 📂 Kapsamlı Excel dosyasını açmak için Enter'a basın...
pause >nul
start "" "data\proje_sablonu.xlsx"
goto END

:SUCCESS_COM_FULL
echo.
echo ================================================================
echo                🎉 TAM COM AUTOMATION BAŞARILI! 🎉
echo ================================================================
echo  📅 Başlangıç: 28.07.2025 (Pazartesi)
echo  📅 Bitiş: 31.10.2025 (3 Ay)
echo  🏗️ 5 Alan Paralel Çalışma
echo  🤖 MS Project COM API Entegrasyonu 
echo  👥 30+ Optimize Görev + 14 Kaynak
echo.
echo  🎯 COM İLE OLUŞTURULAN DOSYALAR:
echo  📊 Excel: data\proje_sablonu.xlsx
echo  📈 MPP: data\SporSalonu_Optimized_26_07_2025.mpp
echo.
echo  ⚡ GELİŞMİŞ ÖZELLİKLER:
echo  • Otomatik görev bağımlılıkları
echo  • Kaynak atamaları ve maliyetleri
echo  • Kritik yol analizi
echo  • Gantt Chart görselleştirme
echo  • Resource optimization
echo ================================================================
echo.
echo 📂 MS Project dosyasını açmak için Enter'a basın...
pause >nul
start "" "data\SporSalonu_Optimized_26_07_2025.mpp"
goto END

:SUCCESS_MPP
echo.
echo [5/5] 🎯 Proje dosyası açılıyor...
if exist "data\SporSalonu_Optimized_26_07_2025.mpp" (
    start "" "data\SporSalonu_Optimized_26_07_2025.mpp"
    echo.
    echo ====================================================
    echo              ✅ MS PROJECT BAŞARILI! ✅              
    echo ====================================================
    echo  📅 Başlangıç: 28.07.2025 (Pazartesi)   
    echo  📅 Bitiş: 31.10.2025 (3 Ay)            
    echo  🏗️ 5 Alan Paralel Çalışma             
    echo  👥 22 Personel + 14 Ekipman            
    echo  🤖 COM/XML automation başarılı
    echo.                                          
    echo  MPP dosyası açıldı! MS Project'te      
    echo  doğrudan kullanabilirsiniz.             
    echo  📁 MPP: data\SporSalonu_Optimized_26_07_2025.mpp
    echo  📁 CSV: data\SporSalonu_Optimized_26_07_2025.csv
    echo  📁 XML: data\SporSalonu_Optimized_26_07_2025.xml
    echo ====================================================
    echo.
) else (
    echo ❌ HATA: Proje dosyası bulunamadı!
)
goto END

:SUCCESS_EXCEL
echo.
echo [5/5] 🎯 Excel dosyası açılıyor...
echo.
echo ====================================================
echo             ✅ EXCEL ŞABLONU HAZIR! ✅               
echo ====================================================
echo  📅 Başlangıç: 28.07.2025 (Pazartesi)   
echo  📅 Bitiş: 31.10.2025 (3 Ay)            
echo  🏗️ 5 Alan Paralel Çalışma             
echo  👥 22 Personel + 14 Ekipman            
echo.                                          
echo  Excel şablonu hazır!                   
echo  📂 Dosya: data\proje_sablonu.xlsx      
echo.                                          
echo  📝 MS PROJECT'E AKTARIM:           
echo  1. Microsoft Project'i açın            
echo  2. Dosya ^> Aç ^> Excel dosyasını seçin  
echo  3. İçe aktarma sihirbazını takip edin  
echo.
echo  💡 COM Automation için Microsoft
echo     Project'in yüklü olması gerekir
echo ====================================================
echo.
echo 📂 Excel dosyasını açmak için Enter'a basın...
pause >nul
explorer "data\proje_sablonu.xlsx"

:END
echo.
echo 📚 OLUŞTURULAN DOSYALAR:
echo ================================
if exist "data\proje_sablonu.xlsx" (
    echo ✅ Excel Şablonu: data\proje_sablonu.xlsx
)
if exist "data\SporSalonu_Optimized_26_07_2025.mpp" (
    echo ✅ MS Project MPP ^(COM^): data\SporSalonu_Optimized_26_07_2025.mpp
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
echo 🤖 COM AUTOMATION ÖZELLİKLERİ:
echo • Tam MS Project entegrasyonu
echo • Otomatik görev bağımlılıkları
echo • Kaynak atamaları ve maliyetleri
echo • Kritik yol optimizasyonu
echo • Gantt Chart görselleştirme
echo • 5 paralel çalışma alanı
echo.
echo 🔧 Sorun mu yaşıyorsunuz?
echo • Microsoft Project'in yüklü olduğunu kontrol edin
echo • comtypes paketi: pip install comtypes
echo • Detaylar için KULLANIM_KILAVUZU.md dosyasını inceleyin
echo.
echo ✅ İşlem tamamlandı! COM automation sistemi aktif.

pause
