#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MS Project COM Automation Sistem Kontrolcüsü
Microsoft Project yüklü mü, COM automation çalışıyor mu kontrol eder
"""

import sys
import os

def check_python():
    """Python versiyon kontrolü"""
    print("🐍 Python kontrol ediliyor...")
    version = sys.version_info
    if version.major >= 3 and version.minor >= 6:
        print(f"   ✅ Python {version.major}.{version.minor}.{version.micro} uygun")
        return True
    else:
        print(f"   ❌ Python {version.major}.{version.minor}.{version.micro} çok eski")
        print("   💡 Python 3.6+ gerekli")
        return False

def check_packages():
    """Gerekli paketleri kontrol et"""
    print("\n📦 Python paketleri kontrol ediliyor...")
    
    packages = ['openpyxl', 'comtypes']
    missing_packages = []
    
    for package in packages:
        try:
            __import__(package)
            print(f"   ✅ {package} paketi mevcut")
        except ImportError:
            print(f"   ❌ {package} paketi eksik")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n   📥 Eksik paketler yükleniyor: {', '.join(missing_packages)}")
        for package in missing_packages:
            try:
                os.system(f"pip install {package}")
                print(f"   ✅ {package} yüklendi")
            except Exception as e:
                print(f"   ❌ {package} yüklenemedi: {e}")
                return False
    
    return True

def check_msproject_installation():
    """Microsoft Project yüklü mü kontrol et"""
    print("\n🏢 Microsoft Project kontrolü...")
    
    try:
        import winreg
        
        # Registry'den MS Project araması
        possible_keys = [
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            r"SOFTWARE\Microsoft\Office\16.0\Common\InstalledPackages",
            r"SOFTWARE\Microsoft\Office\15.0\Common\InstalledPackages",
            r"SOFTWARE\Microsoft\Office\14.0\Common\InstalledPackages",
            r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
        ]
        
        msproject_found = False
        
        for key_path in possible_keys:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                    i = 0
                    while True:
                        try:
                            value_name, value_data, value_type = winreg.EnumValue(key, i)
                            if "project" in value_name.lower() or "project" in str(value_data).lower():
                                print(f"   ✅ MS Project bulundu: {value_name}")
                                msproject_found = True
                                break
                            i += 1
                        except OSError:
                            break
                
                if msproject_found:
                    break
                    
            except FileNotFoundError:
                continue
        
        if not msproject_found:
            # Program Files'da arama
            program_files_paths = [
                r"C:\Program Files\Microsoft Office",
                r"C:\Program Files (x86)\Microsoft Office",
                r"C:\Program Files\Microsoft Office 365",
                r"C:\Program Files (x86)\Microsoft Office 365"
            ]
            
            for path in program_files_paths:
                if os.path.exists(path):
                    for root, dirs, files in os.walk(path):
                        if "winproj.exe" in files or "msproj.exe" in files:
                            print(f"   ✅ MS Project bulundu: {root}")
                            msproject_found = True
                            break
                    if msproject_found:
                        break
        
        if msproject_found:
            print("   ✅ Microsoft Project yüklü görünüyor")
            return True
        else:
            print("   ⚠️ Microsoft Project bulunamadı")
            print("   💡 MS Project'in yüklü olduğundan emin olun")
            return False
            
    except Exception as e:
        print(f"   ❌ MS Project kontrol hatası: {e}")
        return False

def test_com_automation():
    """COM automation test et"""
    print("\n🤖 COM Automation test ediliyor...")
    
    try:
        import comtypes.client
        
        # MS Project COM nesnesini oluşturmayı dene
        print("   🔄 MS Project COM nesnesi oluşturuluyor...")
        app = comtypes.client.CreateObject("MSProject.Application")
        
        # Temel işlemleri test et
        print("   🔄 Temel COM işlemleri test ediliyor...")
        app.Visible = False
        project = app.Projects.Add()
        project.Title = "COM Test Projesi"
        
        # Test görevi ekle
        task = project.Tasks.Add("Test Görevi")
        task.Duration = "1d"
        
        print("   ✅ COM Automation başarılı!")
        
        # Temizlik
        app.Quit()
        
        return True
        
    except Exception as e:
        print(f"   ❌ COM Automation hatası: {e}")
        print("   💡 Olası nedenler:")
        print("      • Microsoft Project yüklü değil")
        print("      • MS Project lisansı yok")
        print("      • COM automation izinleri eksik")
        print("      • MS Project zaten açık (çoklu instance sorunu)")
        return False

def generate_report():
    """Sistem durumu raporu oluştur"""
    print("\n" + "="*60)
    print("📊 MS PROJECT COM AUTOMATION SİSTEM RAPORU")
    print("="*60)
    
    results = {
        'python': check_python(),
        'packages': check_packages(),
        'msproject': check_msproject_installation(),
        'com_automation': False
    }
    
    # COM automation sadece diğerleri başarılıysa test et
    if results['python'] and results['packages']:
        results['com_automation'] = test_com_automation()
    
    print(f"\n📋 SONUÇLAR:")
    print(f"   🐍 Python: {'✅ BAŞARILI' if results['python'] else '❌ BAŞARISIZ'}")
    print(f"   📦 Paketler: {'✅ BAŞARILI' if results['packages'] else '❌ BAŞARISIZ'}")
    print(f"   🏢 MS Project: {'✅ BAŞARILI' if results['msproject'] else '❌ BAŞARISIZ'}")
    print(f"   🤖 COM Automation: {'✅ BAŞARILI' if results['com_automation'] else '❌ BAŞARISIZ'}")
    
    print(f"\n🎯 GENEL DURUM:")
    if all(results.values()):
        print("   🎉 TAM COM AUTOMATION HAZIR!")
        print("   ✅ Tüm sistemler çalışıyor")
        print("   🚀 HIZLI_PROJE.bat'ı çalıştırabilirsiniz")
        status = "FULL_READY"
    elif results['python'] and results['packages'] and results['msproject']:
        print("   ⚠️ COM automation sorunu var")
        print("   💡 MS Project'i kapatıp tekrar deneyin")
        print("   🔄 Fallback modda çalışabilir")
        status = "PARTIAL_READY"
    elif results['python'] and results['packages']:
        print("   ⚠️ MS Project eksik")
        print("   📝 Sadece Excel şablonu oluşturulabilir")
        print("   💡 MS Project'i yükleyin")
        status = "EXCEL_ONLY"
    else:
        print("   ❌ Kritik bileşenler eksik")
        print("   🔧 Gerekli kurulumları yapın")
        status = "NOT_READY"
    
    print(f"\n💡 ÖNERİLER:")
    if not results['msproject']:
        print("   • Microsoft Project'i yükleyin (Professional sürüm önerili)")
    if not results['com_automation'] and results['msproject']:
        print("   • MS Project'i yönetici olarak çalıştırın")
        print("   • Windows güvenlik ayarlarını kontrol edin")
        print("   • Antivirus COM automation'ı engelliyor olabilir")
    if all(results.values()):
        print("   • Sistem hazır, HIZLI_PROJE.bat'ı çalıştırın!")
    
    return status

def main():
    """Ana işlem"""
    print("🔍 MS PROJECT COM AUTOMATION SİSTEM KONTROLCÜSÜ")
    print("=" * 60)
    
    status = generate_report()
    
    print("\n" + "="*60)
    return status

if __name__ == "__main__":
    try:
        final_status = main()
        
        # Exit kodları
        status_codes = {
            "FULL_READY": 0,      # Tam hazır
            "PARTIAL_READY": 1,   # Kısmi hazır
            "EXCEL_ONLY": 2,      # Sadece Excel
            "NOT_READY": 3        # Hazır değil
        }
        
        sys.exit(status_codes.get(final_status, 3))
        
    except KeyboardInterrupt:
        print("\n❌ Kullanıcı tarafından iptal edildi")
        sys.exit(4)
    except Exception as e:
        print(f"\n❌ Beklenmeyen hata: {e}")
        sys.exit(5)
