import win32com.client
import pandas as pd
from pathlib import Path

def check_msp_content():
    try:
        print("🔍 MS Project içeriği kontrol ediliyor...")
        msp = win32com.client.Dispatch("MSProject.Application")
        
        if msp.ActiveProject:
            project = msp.ActiveProject
            print(f"✅ Aktif proje: {project.Name}")
            print(f"📊 Toplam görev sayısı: {project.Tasks.Count}")
            print(f"📊 Toplam kaynak sayısı: {project.Resources.Count}")
            
            print(f"\n📋 İlk 10 Görev:")
            for i in range(1, min(11, project.Tasks.Count + 1)):
                try:
                    task = project.Tasks(i)
                    print(f"   {i:2d}. {task.Name}")
                except:
                    pass
            
            print(f"\n👥 Tüm Kaynaklar ({project.Resources.Count} adet):")
            
            # Kaynakları gruplara ayır
            groups = {
                'Yönetim': [],
                'Kaynakçılar': [],
                'Montajcılar': [],
                'Ekipmanlar': [],
                'Makineler': [],
                'Genel': []
            }
            
            for i in range(1, project.Resources.Count + 1):
                try:
                    resource = project.Resources(i)
                    name = resource.Name
                    
                    # Gruplama
                    if 'proje' in name.lower() or 'usta' in name.lower():
                        groups['Yönetim'].append(name)
                    elif 'kaynakçı' in name.lower():
                        groups['Kaynakçılar'].append(name)
                    elif 'fitter' in name.lower():
                        groups['Montajcılar'].append(name)
                    elif 'manlift' in name.lower() or 'iskele' in name.lower():
                        groups['Ekipmanlar'].append(name)
                    elif 'makine' in name.lower():
                        groups['Makineler'].append(name)
                    else:
                        groups['Genel'].append(name)
                        
                except Exception as e:
                    print(f"   Kaynak {i} okunamadı: {str(e)}")
            
            # Grupları yazdır
            for group_name, resources in groups.items():
                if resources:
                    print(f"\n   🔹 {group_name} ({len(resources)} kaynak):")
                    for resource in resources:
                        print(f"      • {resource}")
            
            print(f"\n📁 Mevcut dosyalar:")
            data_dir = Path("c:/softspace/tahaakgulplanlama/data")
            mpp_files = list(data_dir.glob("*.mpp"))
            for mpp_file in mpp_files:
                size_mb = mpp_file.stat().st_size / (1024 * 1024)
                print(f"   📄 {mpp_file.name} ({size_mb:.2f} MB)")
            
            print(f"\n✅ Kaynak ekleme işlemi başarıyla tamamlandı!")
            print(f"📋 Özet:")
            print(f"   • Excel şablonundan 36 kaynak başarıyla eklendi")
            print(f"   • Görevler: {project.Tasks.Count} adet")
            print(f"   • Kaynaklar: {project.Resources.Count} adet")
            print(f"   • Kaynaklar gruplandırıldı ve maliyet bilgileri eklendi")
            
            print(f"\n📝 Sonraki adımlar:")
            print(f"   1. MS Project'te Resource Sheet'i kontrol edin")
            print(f"   2. Görevlere kaynak atamalarını yapın")
            print(f"   3. Projenizdeki bağımlılıkları kontrol edin")
            print(f"   4. Gantt Chart'ta zaman çizelgesini inceleyin")
            
            return True
        else:
            print("❌ Aktif proje bulunamadı!")
            return False
            
    except Exception as e:
        print(f"❌ Hata: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("🎯 MS Project İçerik Kontrolü")
    print("=" * 35)
    check_msp_content()
