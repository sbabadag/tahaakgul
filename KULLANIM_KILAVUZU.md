# 📊 Excel'den Microsoft Project'e Aktarım Sistemi

## 🚀 Sistem Özeti
Bu sistem, Excel dosyalarından Microsoft Project (.mpp) dosyaları oluşturmak için geliştirilmiştir. Kullanıcılar Excel'de kolay veri girişi yapabilir ve otomatik olarak profesyonel MS Project dosyaları elde edebilir.

## 📁 Dosya Yapısı
```
c:\softspace\tahaakgulplanlama\
├── data/
│   ├── proje_sablonu.xlsx          # 📋 Excel veri girişi şablonu
│   ├── ExceldenMSP.mpp            # 📊 Excel'den oluşturulan Project dosyası
│   ├── TahaAkgul.mpp              # 📊 CSV'den oluşturulan Project dosyası
│   └── spor_salonu_celik_takviye.csv  # 📋 CSV veri dosyası
├── excel_to_msp.py                # 🔄 Excel → MS Project dönüştürücü
├── tahaakgulmpp.py               # 🔄 Akıllı format algılayıcı (Excel/CSV)
├── create_excel_template.py      # 🏗️ Excel şablonu oluşturucu
└── KULLANIM_KILAVUZU.md          # 📖 Bu dosya
```

## 🛠️ Kurulum ve Hazırlık

### 1. Gereksinimler
- ✅ Microsoft Project (2016 veya üzeri)
- ✅ Python 3.7+
- ✅ Gerekli Python paketleri:
  ```bash
  pip install pandas openpyxl pywin32
  ```

### 2. İlk Kurulum
```bash
# 1. Excel şablonunu oluştur
python create_excel_template.py

# 2. Şablon dosyası oluşturuldu: data/proje_sablonu.xlsx
```

## 📋 Excel Şablonu Kullanımı

### 🗂️ Sayfa Yapısı
Excel dosyasında 3 sayfa bulunur:

#### 1. 📊 **Görevler** Sayfası
| Sütun | Açıklama | Örnek |
|-------|----------|-------|
| Görev Adı | Görevin tam adı | "Çelik Montaj İşlemleri" |
| Süre (Gün) | İş günü olarak süre | 5 |
| Başlangıç Tarihi | YYYY-MM-DD formatında | 2025-07-21 |
| Bağımlı Görevler | Önceki görev numaraları (;) | "1;3" |
| Atanan Kaynaklar | Kaynak adları (;) | "Kaynakçı-1;Fitter-1" |
| Görev Türü | Normal/Milestone/Control | "Normal" |
| Öncelik | Düşük/Orta/Yüksek | "Yüksek" |
| Notlar | Ek açıklamalar | "NDT kontrolleri dahil" |

#### 2. 👥 **Kaynaklar** Sayfası
| Sütun | Açıklama | Örnek |
|-------|----------|-------|
| Kaynak Adı | Kaynağın benzersiz adı | "Kaynakçı-1" |
| Kaynak Türü | İnsan/Ekipman | "İnsan" |
| Maksimum Kullanım (%) | Kullanım oranı | 100 |
| Birim Maliyet | Günlük/saatlik maliyet | 2500 |
| Açıklama | Kaynak hakkında bilgi | "Birinci seviye kaynakçı" |

#### 3. 🏢 **Proje Bilgileri** Sayfası
| Özellik | Değer |
|---------|-------|
| Proje Adı | "Spor Salonu Çelik Takviye İşleri" |
| Proje Yöneticisi | "Taha Akgül" |
| Başlangıç Tarihi | "2025-07-21" |
| Bitiş Tarihi | "2025-10-03" |

## 🔄 Aktarım İşlemleri

### Yöntem 1: Sadece Excel → MS Project
```bash
python excel_to_msp.py
```
**Çıktı:** `data/ExceldenMSP.mpp`

### Yöntem 2: Akıllı Format Algılama
```bash
python tahaakgulmpp.py
```
**Davranış:**
- Excel dosyası varsa → Excel'den okur
- Excel yoksa CSV varsa → CSV'den okur
- İkisi de yoksa → Hata verir

**Çıktı:** `data/TahaAkgul.mpp`

## 📝 Veri Girişi İpuçları

### ✅ Doğru Kullanım
```
Görev Adı: "Salon Kaynak İşlemleri"
Süre: 5
Başlangıç: 2025-07-21
Bağımlılar: "7;8"
Kaynaklar: "Kaynakçı-1;Kaynakçı-2;Kaynak Makinesi-1"
```

### ❌ Hatalı Kullanım
```
Görev Adı: (boş)
Süre: "beş gün"
Başlangıç: "yarın"
Bağımlılar: "önceki görev"
Kaynaklar: (yanlış kaynak adı)
```

### 🔗 Bağımlılık Kuralları
- Görev numaraları 1'den başlar
- Birden fazla bağımlılık için `;` kullanın
- Örnek: `"1;3;5"` = Görev 1, 3 ve 5 bitmeden başlamaz

### 👥 Kaynak Atama Kuralları
- Kaynak adları **Kaynaklar** sayfasındakilerle birebir eşleşmeli
- Birden fazla kaynak için `;` kullanın
- Örnek: `"Kaynakçı-1;Fitter-2;26m Manlift-1"`

## 🎯 Proje Optimizasyonu

### 🚀 Performans İpuçları
- **Eşzamanlı Görevler:** Bağımlılık gerektirmeyen görevleri paralel planlayın
- **Kaynak Dengeleme:** Aynı anda çok fazla kaynağı aynı görevde kullanmayın
- **Kritik Yol:** Uzun süreli görevleri dikkatli planlayın

### 📊 Örneklenen Optimizasyonlar
- 5 farklı çalışma alanı (Salon, Fuaye, Spor Salonları, Localar, Teknik Ofisler)
- 22 personel + 14 ekipman = 60 günde tamamlama
- Eşzamanlı çalışma grupları

## 🔧 Sorun Giderme

### ❌ Sık Karşılaşılan Hatalar

#### 1. "Microsoft Project başlatılamadı"
**Çözüm:**
- MS Project'in yüklü olduğundan emin olun
- Yönetici olarak çalıştırın
- Başka Project dosyası açıksa kapatın

#### 2. "Kaynak bulunamadı"
**Çözüm:**
- Kaynak adlarını **Kaynaklar** sayfasından kontrol edin
- Büyük/küçük harf duyarlılığına dikkat edin
- Ekstra boşlukları temizleyin

#### 3. "Geçersiz tarih formatı"
**Çözüm:**
- YYYY-MM-DD formatını kullanın (2025-07-21)
- Excel tarih hücrelerini "Tarih" formatında ayarlayın

#### 4. "Bağımlılık hatası"
**Çözüm:**
- Görev numaralarının doğru olduğundan emin olun
- Döngüsel bağımlılık oluşturmayın (A→B→A)

### 🛠️ Hata Ayıklama
```bash
# Ayrıntılı hata mesajları için verbose mod
python excel_to_msp.py > debug.log 2>&1
```

## 📈 Gelişmiş Özellikler

### 🏗️ Özel Kaynak Türleri
- **İnsan Kaynağı:** Kaynakçı, Fitter, Yönetici
- **Ekipman:** Manlift, İskele, Kaynak Makinesi
- **Malzeme:** Çelik, Kaynak Çubuğu (ileride eklenebilir)

### 📅 Takvim Yönetimi
- Standart: Pazartesi-Cuma, 08:00-17:00
- Özel tatil günleri tanımlanabilir
- Vardiya sistemleri eklenir

### 💰 Maliyet Takibi
- Kaynak bazlı maliyet hesaplama
- Bütçe kontrolü ve raporlama
- Gerçekleşen vs planlanan maliyet

## 🎓 Eğitim Örnekleri

### 📝 Örnek 1: Basit 3 Görevli Proje
```
Görev 1: Hazırlık (2 gün, kaynak: Proje Yöneticisi)
Görev 2: Uygulama (5 gün, bağımlı: 1, kaynak: Kaynakçı-1;Fitter-1)
Görev 3: Kontrol (1 gün, bağımlı: 2, kaynak: Usta Başı)
```

### 📝 Örnek 2: Paralel Çalışma
```
Görev 1: Hazırlık (1 gün)
Görev 2: Alan A İşleri (3 gün, bağımlı: 1, kaynak: Ekip-A)
Görev 3: Alan B İşleri (3 gün, bağımlı: 1, kaynak: Ekip-B)
Görev 4: Genel Bitiş (1 gün, bağımlı: 2;3)
```

## 📞 Destek ve İletişim

### 🐛 Hata Bildirimi
- Hata mesajının tam metnini kaydedin
- Excel dosyasının kopyasını saklayın
- Sistem bilgilerini not alın (Windows sürümü, MS Project sürümü)

### 💡 Özellik İstekleri
- Yeni kaynak türleri
- Farklı takvim sistemleri
- Gelişmiş raporlama

### 📚 Ek Kaynaklar
- Microsoft Project API dokümantasyonu
- Python pandas kılavuzu
- Excel formül referansları

---

## ✅ Hızlı Başlangıç Kontrol Listesi

- [ ] 1. Python ve gerekli paketler yüklü
- [ ] 2. Microsoft Project yüklü ve çalışır durumda
- [ ] 3. Excel şablonu oluşturuldu (`create_excel_template.py`)
- [ ] 4. Proje verileri Excel'e girildi
- [ ] 5. Aktarım scripti çalıştırıldı (`excel_to_msp.py`)
- [ ] 6. MS Project dosyası oluşturuldu ve kontrol edildi

## 🎉 Başarı!
Artık Excel'den Microsoft Project'e profesyonel aktarım yapabilirsiniz!

**Son Güncelleme:** Temmuz 2025 - Versiyon 1.0
**Geliştirici:** Taha Akgül Proje Ekibi
