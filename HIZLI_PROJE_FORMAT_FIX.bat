@echo off
echo.
echo ========================================
echo   SPOR SALONU PROJESI - MS PROJECT      
echo   Format Uyumlu Versiyon                
echo ========================================
echo.

cd /d "%~dp0"

echo [1/1] MS Project uyumlu dosyalar olusturuluyor...
python create_compatible_msp.py

echo.
echo ========================================
echo   ISLEM TAMAMLANDI!
echo ========================================
echo.
echo MS PROJECT'TE ACMAK ICIN:
echo 1. Microsoft Project'i acin
echo 2. Dosya ^> Ac ^> Tur: 'XML Files (*.xml)'
echo 3. data/SporSalonu_MSProject_Compatible.xml dosyasini secin
echo 4. Import Wizard'da 'New Map' secin ve tamamlayin
echo 5. Dosya ^> Farkli Kaydet ^> Tur: 'Project (*.mpp)'
echo.
echo data/ klasorunde uyumlu dosyalar olusturuldu:
echo - SporSalonu_MSProject_Compatible.csv
echo - SporSalonu_MSProject_Compatible.xml
echo.
pause
