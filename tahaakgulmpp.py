import csv, win32com.client as win32
from pathlib import Path
import sys
from datetime import datetime

CSV_FILE   = Path(r"C:\softspace\tahaakgulplanlama\data\plan.csv")
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
except Exception as e:
    print(f"Hata: Microsoft Project başlatılamadı: {e}")
    print("Microsoft Project'in yüklü olduğundan emin olun.")
    sys.exit(1)

# Kaynak havuzu (Resource Sheet)
try:
    with CSV_FILE.open(newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if 'ResourceNames' not in row:
                print("Hata: CSV'de 'ResourceNames' sütunu bulunamadı")
                sys.exit(1)
            for r in row['ResourceNames'].split(';'):
                r = r.strip()
                if r:
                    # Kaynağın zaten var olup olmadığını kontrol et
                    resource_exists = False
                    try:
                        proj.Resources.Item(r)
                        resource_exists = True
                    except:
                        resource_exists = False
                    
                    if not resource_exists:
                        proj.Resources.Add(r)
except FileNotFoundError:
    print(f"Hata: CSV dosyası bulunamadı: {CSV_FILE}")
    sys.exit(1)
except UnicodeDecodeError:
    print("Hata: CSV dosyası UTF-8 kodlaması ile okunamadı")
    sys.exit(1)
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
            
            # Önce kaynak
            if row.get('ResourceNames'):
                for r in row['ResourceNames'].split(';'):
                    r = r.strip()
                    if r:
                        try:
                            resource = proj.Resources.Item(r)
                            # Doğru kaynak atama yöntemi
                            assignment = proj.Assignments.Add(TaskID=t.ID, ResourceID=resource.ID)
                        except Exception as e:
                            print(f"Uyarı: Kaynak '{r}' görev '{t.Name}' için atanamadı: {e}")
            
            # Sonra ilişki
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
