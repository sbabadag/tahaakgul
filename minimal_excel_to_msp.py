import pandas as pd
import win32com.client
from pathlib import Path
from datetime import datetime

def minimal_excel_to_mpp():
    """En basit Excel'den MS Project aktarımı - sadece görevler"""
    
    EXCEL_FILE = Path(r"C:\softspace\tahaakgulplanlama\data\proje_sablonu.xlsx")
    MPP_OUTPUT = Path(r"C:\softspace\tahaakgulplanlama\data\SadeceGorevler.mpp")
    
    print("🚀 Minimal Excel'den MS Project Aktarımı")
    print("=" * 50)
    
    # Excel verilerini oku
    print("📖 Excel görevleri okunuyor...")
    tasks_df = pd.read_excel(EXCEL_FILE, sheet_name='Görevler')
    print(f"   ✓ {len(tasks_df)} görev okundu")
    
    # Microsoft Project'i başlat
    print("🎯 Microsoft Project başlatılıyor...")
    ms_project = win32com.client.Dispatch("MSProject.Application")
    ms_project.Visible = True
    
    # Yeni proje oluştur
    proj = ms_project.Projects.Add()
    proj.Title = "Spor Salonu Çelik Takviye İşleri - Sadece Görevler"
    print("   ✓ Yeni proje oluşturuldu")
    
    # SADECE GÖREVLERİ EKLE
    print(f"📋 {len(tasks_df)} görev ekleniyor...")
    
    for index, task_row in tasks_df.iterrows():
        try:
            task = proj.Tasks.Add()
            task.Name = task_row['Görev Adı']
            
            # Süre ayarla (ana görevler hariç)
            if not task_row['Görev Adı'].isupper():  # Ana görevler büyük harfle yazılmış
                duration = task_row['Süre (Gün)']
                if pd.notna(duration) and duration > 0:
                    task.Duration = f"{int(duration)}d"
                else:
                    task.Duration = "1d"
            
            # Başlangıç tarihi
            if pd.notna(task_row['Başlangıç Tarihi']):
                try:
                    if isinstance(task_row['Başlangıç Tarihi'], str):
                        date_obj = datetime.strptime(task_row['Başlangıç Tarihi'], '%Y-%m-%d')
                        task.Start = date_obj.strftime('%m/%d/%Y')
                except:
                    pass
            
            # Öncelik
            if pd.notna(task_row['Öncelik']):
                priority_map = {'Düşük': 100, 'Orta': 500, 'Yüksek': 1000}
                task.Priority = priority_map.get(task_row['Öncelik'], 500)
            
            # Notlar
            if pd.notna(task_row['Notlar']):
                task.Notes = task_row['Notlar']
            
            # Görev tipini kontrol et
            is_main = task_row['Görev Adı'].isupper()
            task_type = "📁 Ana Görev" if is_main else "📝 Alt Görev"
            print(f"   {index+1:2d}. ✓ {task_type}: {task.Name}")
            
        except Exception as e:
            print(f"   {index+1:2d}. ❌ {task_row['Görev Adı']}: {e}")
    
    print(f"   ✅ Görevler eklendi")
    
    # Dosyayı kaydet
    print("💾 Proje dosyası kaydediliyor...")
    
    try:
        if MPP_OUTPUT.exists():
            MPP_OUTPUT.unlink()
        
        proj.SaveAs(str(MPP_OUTPUT))
        print(f"   ✅ Dosya kaydedildi: {MPP_OUTPUT}")
        
        print("\n🎉 BAŞARILI!")
        print("=" * 50)
        print(f"📁 Dosya: {MPP_OUTPUT}")
        print(f"📊 Görev sayısı: {len(tasks_df)}")
        print("\n📝 Sonraki adımlar:")
        print("   1. MS Project'i açın")
        print("   2. Görevleri kontrol edin")
        print("   3. Kaynakları manuel ekleyin")
        print("   4. Kaynak atamalarını yapın")
        print("   5. Bağımlılıkları ayarlayın")
        
        return True
        
    except Exception as e:
        print(f"❌ Kaydetme hatası: {e}")
        return False

if __name__ == "__main__":
    try:
        minimal_excel_to_mpp()
        input("\nTamamlandı! Enter'a basın...")
    except Exception as e:
        print(f"❌ Hata: {e}")
        input("Enter'a basın...")
