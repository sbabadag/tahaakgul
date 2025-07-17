import pandas as pd
from pathlib import Path

# Excel şablonu oluşturmak için 5 ana mahal ve alt görevler
excel_template_data = {
    'Görev Adı': [
        # 1. SALON ALANI
        'SALON ALANI',
        '  1.1 - Salon Alanında Düşey Çaprazların Montajı',
        '  1.2 - Salon Alanında Düşey Çaprazların İmalatı',
        '  1.3 - Salon Alanında Yatay Çaprazların Eksiklerinin İmalatı',
        '  1.4 - Salon Alanında Yatay Çaprazların Montajı',
        '  1.5 - Salon Alanında Makas ve Ek Platinaların Montajı',
        '  1.6 - Salon Alanında Makas ve Ek Platinaların İmalatı',
        '  1.7 - Salon Alanında Aşık Mahmuz Levhalarının Montajı',
        '  1.8 - Salon Alanında Aşık Mahmuz Levhalarının İmalatı',
        '  1.9 - Salon Alanında Gergi Çubukların Montajı',
        '  1.10 - Salon Alanında Gergi Çubukların İmalatı',
        '  1.11 - Salon Alanında Ankraj Levhalarının Montajı',
        '  1.12 - Salon Alanında Ankraj Levhalarının İmalatı',
        '  1.13 - Salon Alanında Cephe Sistemi Tamirlerinin Yapılması',
        '  1.14 - Salon Alanında Ana Makas Kaynak Eklerinin Tamamlanması',
        
        # 2. FUAYE ALANI
        'FUAYE ALANI',
        '  2.1 - Fuaye Alanında Düşey Çaprazların Montajı',
        '  2.2 - Fuaye Alanında Düşey Çaprazların İmalatı',
        '  2.3 - Fuaye Alanında Yatay Çaprazların Eksiklerinin İmalatı',
        '  2.4 - Fuaye Alanında Yatay Çaprazların Montajı',
        '  2.5 - Fuaye Alanında Makas ve Ek Platinaların Montajı',
        '  2.6 - Fuaye Alanında Makas ve Ek Platinaların İmalatı',
        '  2.7 - Fuaye Alanında Aşık Mahmuz Levhalarının Montajı',
        '  2.8 - Fuaye Alanında Aşık Mahmuz Levhalarının İmalatı',
        '  2.9 - Fuaye Alanında Gergi Çubukların Montajı',
        '  2.10 - Fuaye Alanında Gergi Çubukların İmalatı',
        '  2.11 - Fuaye Alanında Ankraj Levhalarının Montajı',
        '  2.12 - Fuaye Alanında Ankraj Levhalarının İmalatı',
        '  2.13 - Fuaye Alanında Cephe Sistemi Tamirlerinin Yapılması',
        '  2.14 - Fuaye Alanında Ana Makas Kaynak Eklerinin Tamamlanması',
        
        # 3. SPOR SALONLARI
        'SPOR SALONLARI',
        '  3.1 - Spor Salonlarında Düşey Çaprazların Montajı',
        '  3.2 - Spor Salonlarında Düşey Çaprazların İmalatı',
        '  3.3 - Spor Salonlarında Yatay Çaprazların Eksiklerinin İmalatı',
        '  3.4 - Spor Salonlarında Yatay Çaprazların Montajı',
        '  3.5 - Spor Salonlarında Makas ve Ek Platinaların Montajı',
        '  3.6 - Spor Salonlarında Makas ve Ek Platinaların İmalatı',
        '  3.7 - Spor Salonlarında Aşık Mahmuz Levhalarının Montajı',
        '  3.8 - Spor Salonlarında Aşık Mahmuz Levhalarının İmalatı',
        '  3.9 - Spor Salonlarında Gergi Çubukların Montajı',
        '  3.10 - Spor Salonlarında Gergi Çubukların İmalatı',
        '  3.11 - Spor Salonlarında Ankraj Levhalarının Montajı',
        '  3.12 - Spor Salonlarında Ankraj Levhalarının İmalatı',
        '  3.13 - Spor Salonlarında Cephe Sistemi Tamirlerinin Yapılması',
        '  3.14 - Spor Salonlarında Ana Makas Kaynak Eklerinin Tamamlanması',
        
        # 4. LOCALAR
        'LOCALAR',
        '  4.1 - Localarda Düşey Çaprazların Montajı',
        '  4.2 - Localarda Düşey Çaprazların İmalatı',
        '  4.3 - Localarda Yatay Çaprazların Eksiklerinin İmalatı',
        '  4.4 - Localarda Yatay Çaprazların Montajı',
        '  4.5 - Localarda Makas ve Ek Platinaların Montajı',
        '  4.6 - Localarda Makas ve Ek Platinaların İmalatı',
        '  4.7 - Localarda Aşık Mahmuz Levhalarının Montajı',
        '  4.8 - Localarda Aşık Mahmuz Levhalarının İmalatı',
        '  4.9 - Localarda Gergi Çubukların Montajı',
        '  4.10 - Localarda Gergi Çubukların İmalatı',
        '  4.11 - Localarda Ankraj Levhalarının Montajı',
        '  4.12 - Localarda Ankraj Levhalarının İmalatı',
        '  4.13 - Localarda Cephe Sistemi Tamirlerinin Yapılması',
        '  4.14 - Localarda Ana Makas Kaynak Eklerinin Tamamlanması',
        
        # 5. TEKNİK OFİSLER
        'TEKNİK OFİSLER',
        '  5.1 - Teknik Ofislerde Düşey Çaprazların Montajı',
        '  5.2 - Teknik Ofislerde Düşey Çaprazların İmalatı',
        '  5.3 - Teknik Ofislerde Yatay Çaprazların Eksiklerinin İmalatı',
        '  5.4 - Teknik Ofislerde Yatay Çaprazların Montajı',
        '  5.5 - Teknik Ofislerde Makas ve Ek Platinaların Montajı',
        '  5.6 - Teknik Ofislerde Makas ve Ek Platinaların İmalatı',
        '  5.7 - Teknik Ofislerde Aşık Mahmuz Levhalarının Montajı',
        '  5.8 - Teknik Ofislerde Aşık Mahmuz Levhalarının İmalatı',
        '  5.9 - Teknik Ofislerde Gergi Çubukların Montajı',
        '  5.10 - Teknik Ofislerde Gergi Çubukların İmalatı',
        '  5.11 - Teknik Ofislerde Ankraj Levhalarının Montajı',
        '  5.12 - Teknik Ofislerde Ankraj Levhalarının İmalatı',
        '  5.13 - Teknik Ofislerde Cephe Sistemi Tamirlerinin Yapılması',
        '  5.14 - Teknik Ofislerde Ana Makas Kaynak Eklerinin Tamamlanması',
        
        # 6. ORTAK GÖREVLER (Tüm Proje İçin)
        'ORTAK GÖREVLER',
        '  6.1 - Kediyollarının İmalatı',
        '  6.2 - Kediyollarının Montajı',
        '  6.3 - Cephe Sistemi Tamirlerinin Yapılması',
        '  6.4 - Kalite Dosyalarının Hazırlanması',
        '  6.5 - Eksik Tespit Çalışmalarının Yapılması',
        '  6.6 - Koruma Tedbirlerinin Alınması (Saha yanmaz battaniye, pvc örtü, ISG önlemleri)',
        '  6.7 - Mevcut Kedi Merdivenlerinin Tamir Edilmesi',
        '  6.8 - Kenet Çatı Tamirlerinin Yapılması',
        '  6.9 - Sifonik Sistem Yapılması',
        '  6.10 - Mevcut Oluklarının Tamiri (Kaynak) Yapılması',
        '  6.11 - Mevcut Oluklarının Yüzey Temizliği ve Su Yalıtımının Yapılması',
        '  6.12 - Mevcut Oluklarının Üzerine İç Kaplama ve Yeni Saç Yapılması'
    ],
    'Süre (Gün)': [
        # 1. SALON ALANI (Ana mahal + 14 alt görev = 15 görev)
        42, 3, 5, 4, 3, 6, 4, 2, 4, 3, 2, 3, 2, 2, 3,
        
        # 2. FUAYE ALANI (Ana mahal + 14 alt görev = 15 görev)
        35, 3, 4, 3, 3, 5, 3, 2, 3, 2, 2, 2, 2, 2, 3,
        
        # 3. SPOR SALONLARI (Ana mahal + 14 alt görev = 15 görev)
        45, 4, 5, 4, 4, 6, 4, 3, 4, 3, 3, 3, 3, 3, 4,
        
        # 4. LOCALAR (Ana mahal + 14 alt görev = 15 görev)
        28, 2, 3, 2, 2, 4, 2, 1, 2, 2, 1, 2, 1, 1, 2,
        
        # 5. TEKNİK OFİSLER (Ana mahal + 14 alt görev = 15 görev)
        30, 2, 3, 3, 2, 4, 3, 1, 3, 2, 2, 2, 2, 2, 3,
        
        # 6. ORTAK GÖREVLER (Ana mahal + 12 alt görev = 13 görev)
        25, 3, 2, 4, 3, 2, 3, 3, 2, 4, 3, 2, 4
    ],
    'Başlangıç Tarihi': [
        # 1. SALON ALANI (Ana mahal + 14 alt görev = 15 görev)
        '2024-02-01', '2024-02-01', '2024-02-04', '2024-02-09', '2024-02-13', '2024-02-16', '2024-02-22', '2024-02-26', '2024-02-28', '2024-03-04', '2024-03-07', '2024-03-09', '2024-03-12', '2024-03-14', '2024-03-16',
        
        # 2. FUAYE ALANI (Ana mahal + 14 alt görev = 15 görev)
        '2024-02-05', '2024-02-05', '2024-02-08', '2024-02-12', '2024-02-15', '2024-02-18', '2024-02-23', '2024-02-26', '2024-02-28', '2024-03-02', '2024-03-04', '2024-03-06', '2024-03-08', '2024-03-10', '2024-03-12',
        
        # 3. SPOR SALONLARI (Ana mahal + 14 alt görev = 15 görev)
        '2024-02-10', '2024-02-10', '2024-02-14', '2024-02-19', '2024-02-23', '2024-02-27', '2024-03-05', '2024-03-09', '2024-03-12', '2024-03-15', '2024-03-18', '2024-03-21', '2024-03-24', '2024-03-27', '2024-03-30',
        
        # 4. LOCALAR (Ana mahal + 14 alt görev = 15 görev)
        '2024-02-15', '2024-02-15', '2024-02-17', '2024-02-20', '2024-02-22', '2024-02-24', '2024-02-28', '2024-03-01', '2024-03-02', '2024-03-04', '2024-03-05', '2024-03-06', '2024-03-07', '2024-03-08', '2024-03-09',
        
        # 5. TEKNİK OFİSLER (Ana mahal + 14 alt görev = 15 görev)
        '2024-02-20', '2024-02-20', '2024-02-22', '2024-02-25', '2024-02-28', '2024-03-02', '2024-03-06', '2024-03-09', '2024-03-10', '2024-03-12', '2024-03-14', '2024-03-16', '2024-03-18', '2024-03-20', '2024-03-22',
        
        # 6. ORTAK GÖREVLER (Ana mahal + 12 alt görev = 13 görev)
        '2024-03-25', '2024-03-25', '2024-03-28', '2024-03-30', '2024-04-03', '2024-04-06', '2024-04-08', '2024-04-11', '2024-04-14', '2024-04-16', '2024-04-20', '2024-04-23', '2024-04-25'
    ],
    'Bağımlı Görevler': [
        # 1. SALON ALANI (Ana mahal + 14 alt görev = 15 görev)
        '', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14',
        
        # 2. FUAYE ALANI (Ana mahal + 14 alt görev = 15 görev) 
        '', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29',
        
        # 3. SPOR SALONLARI (Ana mahal + 14 alt görev = 15 görev)
        '', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44',
        
        # 4. LOCALAR (Ana mahal + 14 alt görev = 15 görev)
        '', '46', '47', '48', '49', '50', '51', '52', '53', '54', '55', '56', '57', '58', '59',
        
        # 5. TEKNİK OFİSLER (Ana mahal + 14 alt görev = 15 görev)
        '', '61', '62', '63', '64', '65', '66', '67', '68', '69', '70', '71', '72', '73', '74',
        
        # 6. ORTAK GÖREVLER (Ana mahal + 12 alt görev = 13 görev)
        '', '76', '77', '78', '79', '80', '81', '82', '83', '84', '85', '86', '87'
    ],
    'Atanan Kaynaklar': [
        # 1. SALON ALANI (Ana mahal + 14 alt görev = 15 görev)
        'Proje Takımı', 'Kaynakçı[2]', 'Montajcı[3]', 'Kaynakçı[2]', 'Montajcı[2]', 'Kaynakçı[3],Montajcı[2]', 'Kaynakçı[2],Montajcı[1]', 'Montajcı[2]', 'Kaynakçı[2],Montajcı[1]', 'Montajcı[2]', 'Kaynakçı[1]', 'Kaynakçı[2],Montajcı[1]', 'Kaynakçı[1]', 'Montajcı[1]', 'Teknisyen[2]',
        
        # 2. FUAYE ALANI (Ana mahal + 14 alt görev = 15 görev)
        'Proje Takımı', 'Kaynakçı[2]', 'Montajcı[2]', 'Kaynakçı[1]', 'Montajcı[2]', 'Kaynakçı[2],Montajcı[2]', 'Kaynakçı[1],Montajcı[1]', 'Montajcı[1]', 'Kaynakçı[1],Montajcı[1]', 'Montajcı[1]', 'Kaynakçı[1]', 'Kaynakçı[1],Montajcı[1]', 'Kaynakçı[1]', 'Montajcı[1]', 'Teknisyen[1]',
        
        # 3. SPOR SALONLARI (Ana mahal + 14 alt görev = 15 görev)
        'Proje Takımı', 'Kaynakçı[3]', 'Montajcı[4]', 'Kaynakçı[2]', 'Montajcı[3]', 'Kaynakçı[3],Montajcı[3]', 'Kaynakçı[2],Montajcı[2]', 'Montajcı[3]', 'Kaynakçı[2],Montajcı[2]', 'Montajcı[2]', 'Kaynakçı[2]', 'Kaynakçı[2],Montajcı[2]', 'Kaynakçı[2]', 'Montajcı[2]', 'Teknisyen[2]',
        
        # 4. LOCALAR (Ana mahal + 14 alt görev = 15 görev)
        'Proje Takımı', 'Kaynakçı[1]', 'Montajcı[2]', 'Kaynakçı[1]', 'Montajcı[1]', 'Kaynakçı[2],Montajcı[1]', 'Kaynakçı[1],Montajcı[1]', 'Montajcı[1]', 'Kaynakçı[1],Montajcı[1]', 'Montajcı[1]', 'Kaynakçı[1]', 'Kaynakçı[1],Montajcı[1]', 'Kaynakçı[1]', 'Montajcı[1]', 'Teknisyen[1]',
        
        # 5. TEKNİK OFİSLER (Ana mahal + 14 alt görev = 15 görev)
        'Proje Takımı', 'Kaynakçı[1]', 'Montajcı[2]', 'Kaynakçı[1]', 'Montajcı[1]', 'Kaynakçı[2],Montajcı[1]', 'Kaynakçı[1],Montajcı[1]', 'Montajcı[1]', 'Kaynakçı[1],Montajcı[1]', 'Montajcı[1]', 'Kaynakçı[1]', 'Kaynakçı[1],Montajcı[1]', 'Kaynakçı[1]', 'Montajcı[1]', 'Teknisyen[1]',
        
        # 6. ORTAK GÖREVLER (Ana mahal + 12 alt görev = 13 görev)
        'Proje Takımı', 'Kaynakçı[2],Montajcı[2]', 'Montajcı[3]', 'Teknisyen[2]', 'Proje Yöneticisi[1]', 'Teknisyen[1]', 'Montajcı[2],Teknisyen[1]', 'Montajcı[2]', 'Kaynakçı[1],Montajcı[1]', 'Kaynakçı[2]', 'Teknisyen[2]', 'Teknisyen[1]', 'Kaynakçı[1],Montajcı[2]'
    ],
    'Görev Türü': [
        # 1. SALON ALANI (Ana mahal + 14 alt görev = 15 görev) - Ana mahal Summary, alt görevler Normal
        'Summary', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal',
        
        # 2. FUAYE ALANI (Ana mahal + 14 alt görev = 15 görev) - Ana mahal Summary, alt görevler Normal
        'Summary', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal',
        
        # 3. SPOR SALONLARI (Ana mahal + 14 alt görev = 15 görev) - Ana mahal Summary, alt görevler Normal
        'Summary', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal',
        
        # 4. LOCALAR (Ana mahal + 14 alt görev = 15 görev) - Ana mahal Summary, alt görevler Normal
        'Summary', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal',
        
        # 5. TEKNİK OFİSLER (Ana mahal + 14 alt görev = 15 görev) - Ana mahal Summary, alt görevler Normal
        'Summary', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal',
        
        # 6. ORTAK GÖREVLER (Ana mahal + 12 alt görev = 13 görev) - Ana mahal Summary, alt görevler Normal
        'Summary', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal'
    ],
    'Öncelik': [
        # 1. SALON ALANI (Ana mahal + 14 alt görev = 15 görev)
        'Yüksek', 'Yüksek', 'Yüksek', 'Yüksek', 'Orta', 'Yüksek', 'Orta', 'Orta', 'Yüksek', 'Orta', 'Orta', 'Yüksek', 'Orta', 'Orta', 'Orta',
        
        # 2. FUAYE ALANI (Ana mahal + 14 alt görev = 15 görev)
        'Yüksek', 'Yüksek', 'Yüksek', 'Orta', 'Orta', 'Yüksek', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta',
        
        # 3. SPOR SALONLARI (Ana mahal + 14 alt görev = 15 görev)
        'Yüksek', 'Yüksek', 'Yüksek', 'Yüksek', 'Yüksek', 'Yüksek', 'Yüksek', 'Orta', 'Yüksek', 'Orta', 'Orta', 'Yüksek', 'Orta', 'Orta', 'Orta',
        
        # 4. LOCALAR (Ana mahal + 14 alt görev = 15 görev)
        'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta',
        
        # 5. TEKNİK OFİSLER (Ana mahal + 14 alt görev = 15 görev)
        'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta', 'Orta',
        
        # 6. ORTAK GÖREVLER (Ana mahal + 12 alt görev = 13 görev)
        'Yüksek', 'Orta', 'Orta', 'Yüksek', 'Yüksek', 'Yüksek', 'Orta', 'Orta', 'Orta', 'Yüksek', 'Yüksek', 'Orta', 'Yüksek'
    ],
    'Notlar': [
        # 1. SALON ALANI (Ana mahal + 14 alt görev = 15 görev)
        'Salon alanı ana görevleri', 'Salon alanında kaynak yapılması', 'Salon alanında makine yedek parçalarının yapılması', 'Salon alanında yatay çaprazların eksiklerinin imalatı', 'Salon alanında yatay çaprazların montajı', 'Salon alanında makas ve ek platinaların montajı', 'Salon alanında makas ve ek platinaların imalatı', 'Salon alanında aşık mahmuz levhalarının montajı', 'Salon alanında aşık mahmuz levhalarının imalatı', 'Salon alanında gergi çubukların montajı', 'Salon alanında gergi çubukların imalatı', 'Salon alanında ankraj levhalarının montajı', 'Salon alanında ankraj levhalarının imalatı', 'Salon alanında cephe sistemi tamirlerinin yapılması', 'Salon alanında oluk tamiri',
        
        # 2. FUAYE ALANI (Ana mahal + 14 alt görev = 15 görev)
        'Fuaye alanı ana görevleri', 'Fuaye alanında kaynak yapılması', 'Fuaye alanında makine yedek parçalarının yapılması', 'Fuaye alanında yatay çaprazların eksiklerinin imalatı', 'Fuaye alanında yatay çaprazların montajı', 'Fuaye alanında makas ve ek platinaların montajı', 'Fuaye alanında makas ve ek platinaların imalatı', 'Fuaye alanında aşık mahmuz levhalarının montajı', 'Fuaye alanında aşık mahmuz levhalarının imalatı', 'Fuaye alanında gergi çubukların montajı', 'Fuaye alanında gergi çubukların imalatı', 'Fuaye alanında ankraj levhalarının montajı', 'Fuaye alanında ankraj levhalarının imalatı', 'Fuaye alanında cephe sistemi tamirlerinin yapılması', 'Fuaye alanında oluk tamiri',
        
        # 3. SPOR SALONLARI (Ana mahal + 14 alt görev = 15 görev)
        'Spor salonları ana görevleri', 'Spor salonlarında kaynak yapılması', 'Spor salonlarında makine yedek parçalarının yapılması', 'Spor salonlarında yatay çaprazların eksiklerinin imalatı', 'Spor salonlarında yatay çaprazların montajı', 'Spor salonlarında makas ve ek platinaların montajı', 'Spor salonlarında makas ve ek platinaların imalatı', 'Spor salonlarında aşık mahmuz levhalarının montajı', 'Spor salonlarında aşık mahmuz levhalarının imalatı', 'Spor salonlarında gergi çubukların montajı', 'Spor salonlarında gergi çubukların imalatı', 'Spor salonlarında ankraj levhalarının montajı', 'Spor salonlarında ankraj levhalarının imalatı', 'Spor salonlarında cephe sistemi tamirlerinin yapılması', 'Spor salonlarında oluk tamiri',
        
        # 4. LOCALAR (Ana mahal + 14 alt görev = 15 görev)
        'Localar ana görevleri', 'Localarda kaynak yapılması', 'Localarda makine yedek parçalarının yapılması', 'Localarda yatay çaprazların eksiklerinin imalatı', 'Localarda yatay çaprazların montajı', 'Localarda makas ve ek platinaların montajı', 'Localarda makas ve ek platinaların imalatı', 'Localarda aşık mahmuz levhalarının montajı', 'Localarda aşık mahmuz levhalarının imalatı', 'Localarda gergi çubukların montajı', 'Localarda gergi çubukların imalatı', 'Localarda ankraj levhalarının montajı', 'Localarda ankraj levhalarının imalatı', 'Localarda cephe sistemi tamirlerinin yapılması', 'Localarda oluk tamiri',
        
        # 5. TEKNİK OFİSLER (Ana mahal + 14 alt görev = 15 görev)
        'Teknik ofisler ana görevleri', 'Teknik ofislerde kaynak yapılması', 'Teknik ofislerde makine yedek parçalarının yapılması', 'Teknik ofislerde yatay çaprazların eksiklerinin imalatı', 'Teknik ofislerde yatay çaprazların montajı', 'Teknik ofislerde makas ve ek platinaların montajı', 'Teknik ofislerde makas ve ek platinaların imalatı', 'Teknik ofislerde aşık mahmuz levhalarının montajı', 'Teknik ofislerde aşık mahmuz levhalarının imalatı', 'Teknik ofislerde gergi çubukların montajı', 'Teknik ofislerde gergi çubukların imalatı', 'Teknik ofislerde ankraj levhalarının montajı', 'Teknik ofislerde ankraj levhalarının imalatı', 'Teknik ofislerde cephe sistemi tamirlerinin yapılması', 'Teknik ofislerde oluk tamiri',
        
        # 6. ORTAK GÖREVLER (Ana mahal + 12 alt görev = 13 görev)
        'Ortak görevler ana kategorisi', 'Kedi yollarının imalat işleri', 'Kedi yollarının montaj işleri', 'Cephe sistemi genel tamir işleri', 'Kalite kontrol ve belgelendirme', 'Eksik tespit ve raporlama', 'Saha güvenlik tedbirleri ve ISG önlemleri', 'Mevcut kedi merdivenlerinin onarımı', 'Kenet çatı sistemlerinin tamiri', 'Sifonik drenaj sistem kurulumu', 'Oluk kaynak tamir işleri', 'Oluk yüzey hazırlık ve yalıtım', 'Oluk kaplama ve saç işleri'
    ]
}

# DataFrame oluştur
df = pd.DataFrame(excel_template_data)

# Excel dosyasına kaydet
excel_file_path = Path("c:/softspace/tahaakgulplanlama/data/proje_sablonu.xlsx")
excel_file_path.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
    # Ana görev listesi
    df.to_excel(writer, sheet_name='Görevler', index=False)
    
    # Kaynak listesi için ayrı sayfa
    resources_data = {
        'Kaynak Adı': [
            'Proje Yöneticisi (Mimar)', 'Usta Başı',
            'Kaynakçı-1', 'Kaynakçı-2', 'Kaynakçı-3', 'Kaynakçı-4',
            'Kaynakçı-5', 'Kaynakçı-6', 'Kaynakçı-7', 'Kaynakçı-8',
            'Kaynakçı-9', 'Kaynakçı-10', 'Kaynakçı-11', 'Kaynakçı-12',
            'Kaynakçı-13', 'Kaynakçı-14', 'Kaynakçı-15', 'Kaynakçı-16',
            'Fitter-1', 'Fitter-2', 'Fitter-3', 'Fitter-4',
            '26m Manlift-1', '26m Manlift-2', 'Seyyar İskele',
            'Kaynak Makinesi-1', 'Kaynak Makinesi-2', 'Kaynak Makinesi-3',
            'Kaynak Makinesi-4', 'Kaynak Makinesi-5', 'Kaynak Makinesi-6',
            'Kaynak Makinesi-7', 'Kaynak Makinesi-8', 'Kaynak Makinesi-9',
            'Kaynak Makinesi-10', 'Kaynak Makinesi-11'
        ],
        'Kaynak Türü': [
            'İnsan', 'İnsan',
            'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan',
            'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan', 'İnsan',
            'İnsan', 'İnsan', 'İnsan', 'İnsan',
            'Ekipman', 'Ekipman', 'Ekipman',
            'Ekipman', 'Ekipman', 'Ekipman', 'Ekipman', 'Ekipman', 'Ekipman',
            'Ekipman', 'Ekipman', 'Ekipman', 'Ekipman', 'Ekipman'
        ],
        'Maksimum Kullanım (%)': [100] * 36,
        'Birim Maliyet': [
            5000, 3500,  # Yönetici ve Usta
            2500, 2500, 2500, 2500, 2500, 2500, 2500, 2500,  # Kaynakçılar
            2500, 2500, 2500, 2500, 2500, 2500, 2500, 2500,
            2800, 2800, 2800, 2800,  # Fitterlar
            1500, 1500, 800,  # Manlift ve İskele
            300, 300, 300, 300, 300, 300, 300, 300, 300, 300, 300  # Kaynak makineleri
        ],
        'Açıklama': [
            'Sorumlu mimar', 'Saha usta başı',
            'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı',
            'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı',
            'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı',
            'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı', 'Birinci seviye kaynakçı',
            'Çelik montaj uzmanı', 'Çelik montaj uzmanı', 'Çelik montaj uzmanı', 'Çelik montaj uzmanı',
            '26 metre yükseklik kapasiteli', '26 metre yükseklik kapasiteli', 'Taşınabilir çalışma platformu',
            'MIG/MAG kaynak makinesi', 'MIG/MAG kaynak makinesi', 'MIG/MAG kaynak makinesi',
            'MIG/MAG kaynak makinesi', 'MIG/MAG kaynak makinesi', 'MIG/MAG kaynak makinesi',
            'MIG/MAG kaynak makinesi', 'MIG/MAG kaynak makinesi', 'MIG/MAG kaynak makinesi',
            'MIG/MAG kaynak makinesi', 'MIG/MAG kaynak makinesi'
        ]
    }
    
    resources_df = pd.DataFrame(resources_data)
    resources_df.to_excel(writer, sheet_name='Kaynaklar', index=False)
    
    # Proje bilgileri için ayrı sayfa
    project_info = {
        'Özellik': [
            'Proje Adı', 'Proje Yöneticisi', 'Başlangıç Tarihi', 'Bitiş Tarihi',
            'Toplam Süre', 'Çalışma Günleri', 'Çalışma Saatleri', 'Proje Durumu'
        ],
        'Değer': [
            'Spor Salonu Çelik Takviye İşleri', 'Taha Akgül', '2025-07-21', '2025-10-03',
            '60 İş Günü', 'Pazartesi-Cuma', '08:00-17:00', 'Planlama Aşaması'
        ],
        'Açıklama': [
            'Ana proje başlığı', 'Sorumlu proje yöneticisi', 'İlk görevin başlangıcı', 'Son görevin bitişi',
            'Toplam çalışma süresi', 'Haftalık çalışma günleri', 'Günlük çalışma saatleri', 'Mevcut proje durumu'
        ]
    }
    
    project_df = pd.DataFrame(project_info)
    project_df.to_excel(writer, sheet_name='Proje Bilgileri', index=False)

print(f"✅ Excel şablonu oluşturuldu: {excel_file_path}")
print("\n📊 Şablon içeriği:")
print("   • Görevler sayfası: 6 ana kategori + 87 alt görev = 93 toplam görev")
print("   • Kaynaklar sayfası: 36 kaynak tanımı")
print("   • Proje Bilgileri sayfası: Genel proje ayarları")
print("\n📝 Ana Kategoriler:")
print("   1. SALON ALANI (14 alt görev)")
print("   2. FUAYE ALANI (14 alt görev)")
print("   3. SPOR SALONLARI (14 alt görev)")
print("   4. LOCALAR (14 alt görev)")
print("   5. TEKNİK OFİSLER (14 alt görev)")
print("   6. ORTAK GÖREVLER (12 alt görev - Tüm proje için ortak)")
print("\n📝 Ortak Görevler:")
print("   • Kediyollarının İmalatı")
print("   • Kediyollarının Montajı")
print("   • Cephe Sistemi Tamirlerinin Yapılması")
print("   • Kalite Dosyalarının Hazırlanması")
print("   • Eksik Tespit Çalışmalarının Yapılması")
print("   • Koruma Tedbirlerinin Alınması (ISG Önlemleri)")
print("   • Mevcut Kedi Merdivenlerinin Tamir Edilmesi")
print("   • Kenet Çatı Tamirlerinin Yapılması")
print("   • Sifonik Sistem Yapılması")
print("   • Mevcut Oluklarının Tamiri (Kaynak)")
print("   • Mevcut Oluklarının Yüzey Temizliği ve Su Yalıtımı")
print("   • Mevcut Oluklarının Üzerine İç Kaplama ve Yeni Saç")
print("\n📝 Kullanım talimatları:")
print("   1. Excel dosyasını açın")
print("   2. 'Görevler' sayfasında görevlerinizi düzenleyin")
print("   3. 'Kaynaklar' sayfasında kaynaklarınızı kontrol edin")
print("   4. Ana script ile Excel'den MS Project'e aktarım yapın")
