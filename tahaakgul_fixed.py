import csv
import win32com.client
from pathlib import Path
import sys
from datetime import datetime

# Dosya yolları
CSV_FILE = Path(r"C:\softspace\tahaakgulplanlama\data\spor_salonu_celik_takviye.csv")
MPP_OUTPUT = Path(r"C:\softspace\tahaakgulplanlama\data\TahaAkgul_Fixed.mpp")

# CSV dosyasının varlığını kontrol et
if not CSV_FILE.exists():
    print(f"Hata: CSV dosyası bulunamadı: {CSV_FILE}")
    sys.exit(1)

# Çıktı klasörünü oluştur
MPP_OUTPUT.parent.mkdir(parents=True, exist_ok=True)

# Microsoft Project'i başlat
try:
    print("Microsoft Project başlatılıyor...")
    app = win32com.client.Dispatch("MSProject.Application")
    app.Visible = False  # Başlangıçta gizli
    
    # Yeni proje oluştur
    proj = app.Projects.Add()
    
    # Proje özelliklerini ayarla
    proj.Title = "Spor Salonu Çelik Takviye İşleri"
    proj.Author = "Taha Akgül"
    proj.Comments = "60 gün süreli spor salonu çelik takviye projesi - Eşzamanlı çalışma optimizasyonu"
    
    # Başlangıç tarihini ayarla
    try:
        proj.ProjectStart = "7/21/2025"
    except:
        print("Uyarı: Proje başlangıç tarihi ayarlanamadı, varsayılan tarih kullanılacak.")
    
    print("Microsoft Project başarıyla başlatıldı.")
    
except Exception as e:
    print(f"Hata: Microsoft Project başlatılamadı: {e}")
    print("Olası çözümler:")
    print("1. Microsoft Project'in yüklü olduğundan emin olun")
    print("2. Project'i yönetici olarak çalıştırın")
    print("3. Project'in açık olmadığından emin olun")
    sys.exit(1)

# Kaynakları tanımla
resources_to_add = [
    # İnsan kaynakları (22 kişi)
    "Proje Yöneticisi (Mimar)", "Usta Başı",
    # 16 Kaynakçı
    "Kaynakçı-1", "Kaynakçı-2", "Kaynakçı-3", "Kaynakçı-4", 
    "Kaynakçı-5", "Kaynakçı-6", "Kaynakçı-7", "Kaynakçı-8",
    "Kaynakçı-9", "Kaynakçı-10", "Kaynakçı-11", "Kaynakçı-12",
    "Kaynakçı-13", "Kaynakçı-14", "Kaynakçı-15", "Kaynakçı-16",
    # 4 Fitter
    "Fitter-1", "Fitter-2", "Fitter-3", "Fitter-4",
    
    # Ekipmanlar
    "26m Manlift-1", "26m Manlift-2", "Seyyar İskele",
    # Kaynak makineleri
    "Kaynak Makinesi-1", "Kaynak Makinesi-2", "Kaynak Makinesi-3", 
    "Kaynak Makinesi-4", "Kaynak Makinesi-5", "Kaynak Makinesi-6",
    "Kaynak Makinesi-7", "Kaynak Makinesi-8", "Kaynak Makinesi-9",
    "Kaynak Makinesi-10", "Kaynak Makinesi-11"
]

print(f"Toplam {len(resources_to_add)} kaynak tanımlanacak...")

# Kaynakları ekle
try:
    for i, resource_name in enumerate(resources_to_add, 1):
        try:
            resource = proj.Resources.Add()
            resource.Name = resource_name
            
            # İnsan kaynakları için özel ayarlar
            if any(keyword in resource_name for keyword in ["Kaynakçı", "Fitter", "Proje Yöneticisi", "Usta"]):
                resource.MaxUnits = 100.0  # %100 kullanım
            else:
                # Ekipman için
                resource.MaxUnits = 100.0
                
            print(f"{i:2d}. Kaynak eklendi: {resource_name}")
            
        except Exception as e:
            print(f"Uyarı: Kaynak '{resource_name}' eklenemedi: {e}")
    
    print(f"✓ Kaynaklar başarıyla tanımlandı.")
    
except Exception as e:
    print(f"Hata: Kaynaklar tanımlanırken genel hata: {e}")
    sys.exit(1)

# CSV'den görevleri oku ve ekle
task_objs = []
try:
    print("\nGörevler okunuyor ve ekleniyor...")
    
    with CSV_FILE.open(newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=1):
            try:
                # Görev ekle
                task = proj.Tasks.Add()
                task.Name = row.get('Name', f'Görev {i}')
                
                # Süre ayarla
                duration_str = row.get('Duration', '1')
                try:
                    duration_days = int(duration_str.replace('d', '').strip())
                    task.Duration = f"{duration_days}d"
                except:
                    task.Duration = "1d"
                
                # Başlangıç tarihi ayarla (eğer varsa)
                if row.get('Start'):
                    try:
                        start_date = row['Start'].strip()
                        if start_date and start_date != '':
                            # Tarih formatını çevir
                            date_obj = datetime.strptime(start_date, '%Y-%m-%d')
                            task.Start = date_obj.strftime('%m/%d/%Y')
                    except Exception as date_e:
                        print(f"Uyarı: Tarih ayarlanamadı '{row.get('Start')}' için: {date_e}")
                
                task_objs.append(task)
                print(f"{i:2d}. Görev eklendi: {task.Name}")
                
            except Exception as e:
                print(f"Hata: Görev {i} eklenirken: {e}")
                
    print(f"✓ {len(task_objs)} görev başarıyla eklendi.")
    
except Exception as e:
    print(f"Hata: Görevler eklenirken: {e}")
    sys.exit(1)

# Kaynak atamalarını ve bağımlılıkları ekle
try:
    print("\nKaynak atamaları ve bağımlılıklar ekleniyor...")
    
    with CSV_FILE.open(newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=1):
            if i > len(task_objs):
                break
                
            task = task_objs[i-1]
            
            # Kaynak atamalarını yap
            if row.get('ResourceNames'):
                resource_names = row['ResourceNames'].split(';')
                for resource_name in resource_names:
                    resource_name = resource_name.strip()
                    if resource_name:
                        try:
                            # Kaynağı bul
                            resource_found = None
                            for resource in proj.Resources:
                                if resource.Name == resource_name:
                                    resource_found = resource
                                    break
                            
                            if resource_found:
                                # Basit atama yöntemi
                                try:
                                    assignment = proj.Assignments.Add()
                                    assignment.TaskUniqueID = task.UniqueID
                                    assignment.ResourceUniqueID = resource_found.UniqueID
                                    assignment.Work = task.Work
                                    print(f"    ✓ Atama: {resource_name} -> {task.Name}")
                                except Exception as assign_e:
                                    print(f"    ⚠ Atama hatası: {resource_name} -> {task.Name}: {assign_e}")
                            else:
                                print(f"    ⚠ Kaynak bulunamadı: {resource_name}")
                                
                        except Exception as e:
                            print(f"    ⚠ Kaynak atama hatası: {e}")
            
            # Bağımlılıkları ekle
            if row.get('Predecessors'):
                predecessors = row['Predecessors'].split(';')
                for pred in predecessors:
                    pred = pred.strip()
                    if pred:
                        try:
                            # Sadece sayısal değeri al
                            pred_id = int(pred.replace('FS', '').replace('SS', '').replace('FF', '').replace('SF', '').strip())
                            
                            if 1 <= pred_id <= len(task_objs):
                                pred_task = task_objs[pred_id - 1]
                                task.PredecessorTasks.Add(pred_task)
                                print(f"    ✓ Bağımlılık: {pred_task.Name} -> {task.Name}")
                            else:
                                print(f"    ⚠ Geçersiz predecessor ID: {pred_id}")
                                
                        except Exception as e:
                            print(f"    ⚠ Bağımlılık hatası: {pred}: {e}")

    print("✓ Kaynak atamaları ve bağımlılıklar tamamlandı.")

except Exception as e:
    print(f"Hata: Kaynak atamaları sırasında: {e}")

# Dosyayı kaydet
try:
    print(f"\nProje dosyası kaydediliyor: {MPP_OUTPUT}")
    app.Visible = True  # Kaydetmeden önce görünür yap
    proj.SaveAs(str(MPP_OUTPUT))
    print(f"✅ Proje başarıyla kaydedildi: {MPP_OUTPUT}")
    
    # Özet bilgiler
    print(f"\n📊 PROJE ÖZETİ:")
    print(f"   • Toplam görev sayısı: {len(task_objs)}")
    print(f"   • Toplam kaynak sayısı: {len(resources_to_add)}")
    print(f"   • İnsan kaynağı: 22 kişi (Mimar + Usta + 16 Kaynakçı + 4 Fitter)")
    print(f"   • Ekipman: 14 adet (2 Manlift + 1 İskele + 11 Kaynak Makinesi)")
    print(f"   • Proje süresi: 60 iş günü")
    print(f"   • Başlangıç: 21 Temmuz 2025")
    
except Exception as e:
    print(f"❌ Dosya kaydetme hatası: {e}")
    
finally:
    try:
        # Uygulamayı kapat
        input("\nProje kontrol edildi mi? Enter'a basarak Microsoft Project'i kapatın...")
        app.Quit()
        print("Microsoft Project kapatıldı.")
    except:
        print("Microsoft Project kapatma işlemi tamamlanamadı.")

print("\n🎉 İşlem tamamlandı!")
