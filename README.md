"# Taha Akgül Proje Planlama Aracı

Bu Python scripti, CSV dosyasından Microsoft Project (.mpp) dosyası oluşturur.

## Gereksinimler

1. **Microsoft Project** - Bilgisayarınızda yüklü olmalı
2. **Python paketleri:**
   ```
   pip install pywin32
   ```

## CSV Dosya Formatı

CSV dosyası şu sütunları içermelidir:

- **Name** (zorunlu): Görev adı
- **Duration** (zorunlu): Süre (örn: "5d", "2w", "3h")
- **Start** (opsiyonel): Başlangıç tarihi (YYYY-MM-DD formatında)
- **ResourceNames** (opsiyonel): Kaynak isimleri (noktalı virgülle ayrılmış)
- **Predecessors** (opsiyonel): Önceki görevler (örn: "1FS", "2SS", "3")

## Kullanım

1. `data/plan.csv` dosyasını oluşturun (örnek için `plan_sample.csv`'ye bakın)
2. Scripti çalıştırın:
   ```
   python tahaakgulmpp.py
   ```
3. Oluşturulan `.mpp` dosyası `data/TahaAkgul.mpp` konumunda olacak

## Özellikler

- ✅ Hata kontrolü ve kullanıcı dostu mesajlar
- ✅ Kaynak yönetimi
- ✅ Görev bağımlılıkları (FS, SS, FF, SF)
- ✅ Otomatik klasör oluşturma
- ✅ UTF-8 kodlama desteği

## Predecessor Tipleri

- **FS** (Finish-to-Start): Varsayılan
- **SS** (Start-to-Start)
- **FF** (Finish-to-Finish) 
- **SF** (Start-to-Finish)

Örnek: "2FS" = Görev 2 bittikten sonra başla" 
