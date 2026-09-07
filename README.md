# 🚚 Treyler & Dorse Cari ve Fatura Takip Sistemi

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.110%2B-009688?style=for-the-badge&logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![SQLite](https://img.shields.io/badge/SQLite-3-003B57?style=for-the-badge&logo=sqlite&logoColor=white)](https://www.sqlite.org/)
[![Docker](https://img.shields.io/badge/Docker-Supported-2496ED?style=for-the-badge&logo=docker&logoColor=white)](https://www.docker.com/)
[![License](https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge)](LICENSE)

> **Treyler, dorse ve ağır vasıta imalat/servis sektörü için özel olarak tasarlanmış, doğal dil algılamalı akıllı asistan destekli modern cari ve fatura yönetim sistemi.**

---

## 📖 Genel Bakış

**Treyler & Dorse Cari ve Fatura Takip Sistemi**, karmaşık muhasebe programlarının zorluklarını ortadan kaldırarak; atölye, fabrika ve işletmelerin fatura, tahsilat, ödeme ve bakiye süreçlerini en hızlı ve zahmetsiz şekilde takip etmesini sağlayan hafif ve modern bir web uygulamasıdır.

---

## ✨ Temel Özellikler

### 🧠 1. Akıllı Asistan (Smart Parser)
- Serbest dilde yazılan işlemleri anında anlar ve otomatik form doldurur.
- **Örnek girdiler:**
  - `Özdemir Nakliyat 150000 fatura`
  - `Kaya Lojistik 50 bin havale geldi`
  - `Aksan Dingil 85000 fatura geldi`
  - `Çeliksan 40 bin ödeme yaptık`
- Plaka ve şasi numaralarını tutardan akıllıca ayırt eder.
- Tek tıkla onaylanarak bakiye otomatik güncellenir.

### ⚡ 2. Hızlı ve Renkli İşlem Butonları
- **🔵 Fatura Kes:** Satış faturası girişi (Müşterinin borcu / alacağımız artar).
- **🟢 Tahsilat Al:** Müşteriden gelen havale veya nakit tahsilatı.
- **🟠 Fatura Geldi:** Tedarikçiden veya servisten gelen alış faturası.
- **🔴 Ödeme Yap:** Tedarikçiye yapılan ödeme.
- Dorse tipi (*Damper, Lowbed, Tenteli, Silobas vb.*), plaka ve şasi numarası kaydetme desteği.

### 📊 3. KolaySoft E-Fatura Entegrasyonu
- KolaySoft portalından indirilen fatura Excel tablolarını sürükle-bırak ile tek tıkla içeri aktarma.
- Listede geçen yeni firmaları otomatik tespit edip sisteme ekler.
- Faturaları carilerin hesaplarına toplu olarak hatasız işler.

### 📑 4. Raporlama ve Excel Çıktıları
- **Cari Bakiye Listesi:** Renkli, biçimlendirilmiş ve genel toplamlı Excel tablosu oluşturma.
- **Cari Hesap Ekstresi:** İlgili firmanın tüm geçmiş hareketlerini döküm, yazdırma ve Excel olarak dışa aktarma imkânı.

### 💬 5. WhatsApp Hızlı Bakiye Bildirimi
- Müşterinin cari durumunu ve güncel bakiyesini içeren hazır metin ile tek tıkla WhatsApp Web üzerinden mesaj taslağı oluşturma. *(Onayınız olmadan mesaj gönderilmez)*

### 🔄 6. Güvenli İşlem & Geri Alma (İptal)
- Yanlış girilen veya hatalı tutarlı işlemleri tek tıkla silme desteği.
- Silinen işlem sonrasında cari hesap bakiyesi hatasız olarak anında yeniden dengelenir.

### 💾 7. Veritabanı ve Yedekleme
- Kolay taşınabilir ve kurulum gerektirmeyen SQLite mimarisi (`muhasebe.db`).
- Arayüzden tek tıkla zaman damgalı veritabanı yedeği indirme.

---

## 🛠️ Teknoloji Yığını

| Katman | Teknoloji / Kütüphane | Açıklama |
|---|---|---|
| **Backend** | [FastAPI](https://fastapi.tiangolo.com/) | Yüksek performanslı modern Python web çatısı |
| **Sunucu** | [Uvicorn](https://www.uvicorn.org/) | Asenkron ASGI web sunucusu |
| **Veritabanı** | [SQLite3](https://www.sqlite.org/) | Sıfır konfigürasyonlu dosya tabanlı veritabanı |
| **Excel Motoru** | [openpyxl](https://openpyxl.readthedocs.io/) | Gelişmiş Excel okuma ve biçimlendirilmiş rapor oluşturma |
| **Doğal Dil & Regex** | Dahili `smart_parser.py` | Akıllı metin çözümleme motoru |
| **Konteyner** | [Docker](https://www.docker.com/) | Platform bağımsız hızlı kurulum ve dağıtım |

---

## 📁 Proje Dizin Yapısı

```text
├── BASLAT.bat             # Windows için tek tıkla başlatma betiği
├── Dockerfile             # Docker imajı yapılandırması
├── Procfile               # Bulut dağıtım (PaaS) servisleri için başlatma tanımı
├── requirements.txt       # Python bağımlılık listesi
├── KULLANIM_KILAVUZU.txt  # Detaylı son kullanıcı kılavuzu
│
├── app.py                 # FastAPI ana API ve web sunucu yönlendirmeleri
├── database.py            # SQLite şeması ve veritabanı CRUD işlemleri
├── smart_parser.py        # Akıllı asistanın doğal dil ayrıştırma motoru
├── excel_service.py       # KolaySoft içe aktarım ve Excel raporlama servisi
│
├── templates/
│   └── index.html         # Responsive, modern web arayüzü
│
└── muhasebe.db            # SQLite yerel veritabanı dosyası
```

---

## 🚀 Hızlı Başlangıç ve Kurulum

### Yöntem 1: Windows Tek Tıkla Başlatma (En Kolay)
Bilgisayarınızda Python yüklü ise:
1. Klasördeki **`BASLAT.bat`** dosyasına çift tıklayın.
2. Web tarayıcınız (`http://127.0.0.1:8000`) otomatik olarak açılacaktır.

---

### Yöntem 2: Python ile Manuel Çalıştırma

1. **Depoyu Klonlayın:**
   ```bash
   git clone https://github.com/kullaniciadi/hesap-takip.git
   cd hesap-takip
   ```

2. **Sanal Ortam (Virtualenv) Oluşturun (Önerilir):**
   ```bash
   python -m venv venv
   # Windows için:
   venv\Scripts\activate
   # macOS/Linux için:
   source venv/bin/activate
   ```

3. **Gerekli Paketleri Yükleyin:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Uygulamayı Başlatın:**
   ```bash
   python app.py
   ```
   *Tarayıcınızdan `http://localhost:8000` adresine gidin.*

---

### Yöntem 3: Docker ile Çalıştırma

1. **Docker İmajını Oluşturun:**
   ```bash
   docker build -t treyler-cari-takip .
   ```

2. **Kapsayıcıyı Çalıştırın:**
   ```bash
   docker run -d -p 8000:8000 --name treyler-takip treyler-cari-takip
   ```
   *`http://localhost:8000` adresinden erişebilirsiniz.*

---

## 💡 Kullanım İpuçları

1. **Akıllı Asistan:** Giriş ekranının en üstündeki kutucuğa firma adını ve tutarı doğrudan yazıp `Enter` tuşuna basmanız yeterlidir.
2. **KolaySoft Aktarımı:** Portalınızdan indirdiğiniz `.xlsx` dosyasını sürükleyin; sistem faturaları carilerine göre otomatik eşler ve bakiyeleri günceller.
3. **Yedek Almayı Unutmayın:** Düzenli aralıklarla sağ üst köşedeki **"Yedek Al"** düğmesini kullanarak `muhasebe.db` dosyanızın kopyasını indirebilirsiniz.

---

## 📄 Lisans

Bu proje [MIT](LICENSE) lisansı altında sunulmaktadır.
