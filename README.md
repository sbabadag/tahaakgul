# 🏗️ Spor Salonu Proje Planlama Sistemi

## 🎯 HIZLI BAŞLANGIÇ

### ⚡ Format Uyumlu Çözüm (ÖNERİLEN)
```batch
HIZLI_PROJE_FORMAT_FIX.bat
```

**MS Project'te açmak için:**
1. Microsoft Project'i açın
2. Dosya > Aç > Türü: 'XML Files (*.xml)'
3. `data/SporSalonu_MSProject_Compatible.xml` dosyasını seçin
4. Import Wizard'da 'New Map' seçin ve tamamlayın

## 📊 Proje Özellikleri
- **28 görev** - Detaylı çelik konstrüksiyon planı
- **15 kaynak** - Personel, ekipman, araçlar
- **Tarih**: 28.07.2025 → 31.10.2025 (3 ay)
- **Strateji**: 5 alan paralel çalışma

## 🔧 Alternatif Sistemler
- `HIZLI_PROJE.bat` - COM automation sistemi
- `HIZLI_PROJE_COM.bat` - Gelişmiş COM automation

## 📖 Detaylar
Tüm detaylar için `KULLANIM_KILAVUZU.md` dosyasını inceleyin.

---
**Geliştirici:** Taha Akgül Proje Planlama Sistemi  
**Versiyon:** 3.0 (Temizlenmiş) - Temmuz 2025

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
