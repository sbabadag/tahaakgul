import csv, win32com.client as win32
from pathlib import Path
import sys
from datetime import datetime

CSV_FILE   = Path(r"C:\softspace\tahaakgulplanlama\data\spor_salonu_celik_takviye.csv")
MPP_OUTPUT = Path(r"C:\softspace\tahaakgulplanlama\data\TahaAkgul.mpp")

# CSV dosyasının varlığını kontrol et
if not CSV_FILE.exists():
    print(f"Hata: CSV dosyası bulunamadı: {CSV_FILE}")
    sys.exit(1)

# Çıktı klasörünü oluştur
MPP_OUTPUT.parent.mkdir(parents=True, exist_ok=True)

try:
    # Project'i başlat
    app = win32.Dispatch("MSProject.Application")
    app.Visible = True
    proj = app.Projects.Add()          # yeni boş proje
    
    # Proje özelliklerini ayarla
    proj.Title = "Spor Salonu Çelik Takviye İşleri"
    proj.Author = "Taha Akgül"
    proj.Comments = "60 gün süreli spor salonu çelik takviye projesi - Eşzamanlı çalışma optimizasyonu"
    proj.ProjectStart = "7/21/2025"  # MM/DD/YYYY formatı
    
    print("Proje oluşturuldu ve özellikler ayarlandı.")
except Exception as e:
    print(f"Hata: Microsoft Project başlatılamadı: {e}")
    print("Microsoft Project'in yüklü olduğundan emin olun.")
    sys.exit(1)

# Önce tüm kaynakları tanımla
resources_to_add = [
    # İnsan kaynakları
    "Proje Yöneticisi", "Usta Başı",
    # 16 Kaynakçı
    "Kaynakçı1", "Kaynakçı2", "Kaynakçı3", "Kaynakçı4", "Kaynakçı5", "Kaynakçı6", 
    "Kaynakçı7", "Kaynakçı8", "Kaynakçı9", "Kaynakçı10", "Kaynakçı11", "Kaynakçı12",
    "Kaynakçı13", "Kaynakçı14", "Kaynakçı15", "Kaynakçı16",
    # 17 Fitter/Montajcı 
    "Fitter1", "Fitter2", "Fitter3", "Fitter4", "Fitter5", "Fitter6", "Fitter7",
    "Fitter8", "Fitter9", "Fitter10", "Fitter11", "Fitter12", "Fitter13", "Fitter14",
    "Fitter15", "Fitter16", "Fitter17",
    # Ekipmanlar
    "Manlift1", "Manlift2",  # 2x 26 metre ahtapot manlift
    "Iskele",  # Çatının tamamını kaplayan üstten yürüyen iskele
    # 11 Kaynak makinesi
    "Kaynak Makinesi1", "Kaynak Makinesi2", "Kaynak Makinesi3", "Kaynak Makinesi4",
    "Kaynak Makinesi5", "Kaynak Makinesi6", "Kaynak Makinesi7", "Kaynak Makinesi8",
    "Kaynak Makinesi9", "Kaynak Makinesi10", "Kaynak Makinesi11"
]

print("Kaynaklar tanımlanıyor...")
try:
    for resource_name in resources_to_add:
        try:
            resource = proj.Resources.Add()
            resource.Name = resource_name
            
            # Kaynak türüne göre özellikler belirle
            if any(keyword in resource_name for keyword in ["Kaynakçı", "Fitter", "Proje Yöneticisi", "Usta Başı"]):
                # İnsan kaynağı (Work)
                resource.MaxUnits = 100  # %100 kullanım
            else:
                # Ekipman/Malzeme (Material)
                resource.MaxUnits = 100  # %100 kullanım
                
            print(f"Kaynak eklendi: {resource_name}")
        except Exception as e:
            print(f"Uyarı: Kaynak '{resource_name}' eklenemedi: {e}")
    
    print(f"Toplam {len(resources_to_add)} kaynak tanımlandı.")
except Exception as e:
    print(f"Hata: Kaynaklar eklenirken hata oluştu: {e}")
    sys.exit(1)

# Görevleri ekle — ID sırası CSV'deki sırayla olur
task_objs = []
try:
    with CSV_FILE.open(newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Gerekli sütunları kontrol et
            required_columns = ['Name', 'Duration']
            for col in required_columns:
                if col not in row:
                    print(f"Hata: CSV'de '{col}' sütunu bulunamadı")
                    sys.exit(1)
            
            t = proj.Tasks.Add(row['Name'])
            t.Duration = row['Duration']
            if row.get('Start'):  # Start sütunu opsiyonel
                try:
                    # Tarih formatını Microsoft Project'e uygun hale getir
                    date_str = row['Start'].strip()
                    if date_str:
                        # Microsoft Project için farklı tarih formatlarını dene
                        try:
                            # Önce MM/DD/YYYY formatını dene
                            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                            t.Start = date_obj.strftime('%m/%d/%Y')
                        except:
                            try:
                                # DD/MM/YYYY formatını dene
                                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                                t.Start = date_obj.strftime('%d/%m/%Y')
                            except:
                                # Son çare olarak orijinal formatı dene
                                t.Start = date_str
                except ValueError as e:
                    print(f"Uyarı: Geçersiz tarih formatı '{row['Start']}' görev '{row['Name']}' için. Tarih atlanıyor.")
                except Exception as e:
                    print(f"Uyarı: Tarih atanırken hata oluştu '{row['Start']}' görev '{row['Name']}' için: {e}")
            task_objs.append(t)
except Exception as e:
    print(f"Hata: Görevler eklenirken hata oluştu: {e}")
    sys.exit(1)

# Bağımlılık ve kaynak atamalarını ekle (ikinci turda ID'ler hazır)
try:
    with CSV_FILE.open(newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader, start=1):
            t = task_objs[i-1]
            
            # Önce kaynak ataması
            if row.get('ResourceNames'):
                for r in row['ResourceNames'].split(';'):
                    r = r.strip()
                    if r:
                        try:
                            # Kaynağı bul
                            resource = None
                            for res in proj.Resources:
                                if res.Name == r:
                                    resource = res
                                    break
                            
                            if resource:
                                # Doğru kaynak atama yöntemi
                                assignment = proj.Assignments.Add()
                                assignment.TaskUniqueID = t.UniqueID
                                assignment.ResourceUniqueID = resource.UniqueID
                                print(f"    Kaynak atandı: {r} -> {t.Name}")
                            else:
                                print(f"Uyarı: Kaynak '{r}' bulunamadı")
                        except Exception as e:
                            print(f"Uyarı: Kaynak '{r}' görev '{t.Name}' için atanamadı: {e}")
            
            # Sonra bağımlılık ilişkileri
            if row.get('Predecessors'):
                for pred in row['Predecessors'].split(';'):
                    pred = pred.strip()
                    if pred:
                        try:
                            # FS, SS, FF, SF gibi ilişki tiplerini ayır
                            if pred.endswith('FS'):
                                pid = int(pred.rstrip('FS'))
                            elif pred.endswith('SS'):
                                pid = int(pred.rstrip('SS'))
                            elif pred.endswith('FF'):
                                pid = int(pred.rstrip('FF'))
                            elif pred.endswith('SF'):
                                pid = int(pred.rstrip('SF'))
                            else:
                                pid = int(pred)  # Sadece sayı varsa FS varsay
                            
                            if 1 <= pid <= len(task_objs):
                                t.PredecessorTasks.Add(proj.Tasks.Item(pid))
                            else:
                                print(f"Uyarı: Geçersiz predecessor ID: {pid}")
                        except (ValueError, Exception) as e:
                            print(f"Uyarı: Predecessor '{pred}' işlenemedi: {e}")
except Exception as e:
    print(f"Hata: Bağımlılıklar ve kaynak atamaları eklenirken hata oluştu: {e}")
    sys.exit(1)

# Dosyayı kaydet
try:
    proj.SaveAs(str(MPP_OUTPUT))
    print(f"Başarılı! Proje dosyası oluşturuldu: {MPP_OUTPUT}")
except Exception as e:
    print(f"Hata: Dosya kaydedilirken hata oluştu: {e}")
    sys.exit(1)
finally:
    # Uygulamayı kapat
    try:
        app.Quit()
    except:
        pass
