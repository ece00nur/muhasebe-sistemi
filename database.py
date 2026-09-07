import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "muhasebe.db")

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # Cariler Tablosu (Müşteriler & Tedarikçiler)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS cariler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        unvan TEXT NOT NULL,
        tip TEXT DEFAULT 'Musteri', -- 'Musteri', 'Tedarikci', 'Diger'
        yetkili TEXT DEFAULT '',
        telefon TEXT DEFAULT '',
        eposta TEXT DEFAULT '',
        vergi_no TEXT DEFAULT '',
        vergi_dairesi TEXT DEFAULT '',
        adres TEXT DEFAULT '',
        bakiye REAL DEFAULT 0.0, -- Pozitif: Müşterinin bize borcu (Alacağımız), Negatif: Bizim tedarikçiye borcumuz
        notlar TEXT DEFAULT '',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # İşlemler Tablosu (Faturalar, Tahsilatlar, Ödemeler)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS islemler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cari_id INTEGER NOT NULL,
        tur TEXT NOT NULL, -- 'Satis_Faturasi', 'Alis_Faturasi', 'Tahsilat', 'Odeme'
        belge_no TEXT DEFAULT '',
        tarih TEXT NOT NULL,
        vade_tarihi TEXT DEFAULT '',
        tutar REAL NOT NULL,
        kdv_orani REAL DEFAULT 20.0,
        tevkifat TEXT DEFAULT '', -- Örn: '5/10', 'Yok'
        odeme_yontemi TEXT DEFAULT 'Havale/EFT', -- 'Nakit', 'Havale/EFT', 'Kredi Karti', 'Cek', 'Senet'
        sasi_no TEXT DEFAULT '',
        dorse_tipi TEXT DEFAULT '', -- 'Damper', 'Lowbed', 'Tenteli', 'Konteyner Tasiyici', 'Silobas', 'Kapakli', 'Yedek Parca/Servis', 'Diger'
        plaka TEXT DEFAULT '',
        aciklama TEXT DEFAULT '',
        kolaysoft_id TEXT DEFAULT '', -- Kolaysoft ETTN / Belge No eşleşmesi
        durum TEXT DEFAULT 'Aktif',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (cari_id) REFERENCES cariler(id) ON DELETE CASCADE
    )
    """)
    
    # Çek / Senet Tablosu (Treyler sektöründe çok yaygın)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS cek_senet (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cari_id INTEGER NOT NULL,
        islem_id INTEGER,
        tur TEXT NOT NULL, -- 'Alinan_Cek', 'Verilen_Cek', 'Alinan_Senet', 'Verilen_Senet'
        belge_no TEXT DEFAULT '',
        vade_tarihi TEXT NOT NULL,
        tutar REAL NOT NULL,
        banka TEXT DEFAULT '',
        sube TEXT DEFAULT '',
        kesideci TEXT DEFAULT '',
        durum TEXT DEFAULT 'Portfoyde', -- 'Portfoyde', 'Tahsil_Edildi', 'Cirolandi', 'Karsiliksiz', 'Odendi'
        aciklama TEXT DEFAULT '',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (cari_id) REFERENCES cariler(id) ON DELETE CASCADE
    )
    """)
    
    # Firma Ayarları
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ayarlar (
        anahtar TEXT PRIMARY KEY,
        deger TEXT
    )
    """)
    
    conn.commit()
    
    # Varsayılan Firma Ayarları
    cursor.execute("SELECT COUNT(*) FROM ayarlar")
    if cursor.fetchone()[0] == 0:
        cursor.executemany("INSERT INTO ayarlar (anahtar, deger) VALUES (?, ?)", [
            ("firma_adi", "Treyler & Dorse Sanayi"),
            ("firma_telefon", ""),
            ("firma_yetkili", ""),
            ("varsayilan_kdv", "20"),
            ("whatsapp_sablonu", "Sayın {unvan} yetkilisi, firmamızdaki güncel cari hesap bakiyeniz: {bakiye} ₺ ({durum}) olarak görünmektedir. Detaylı bilgi için bizimle iletişime geçebilirsiniz. Hayırlı işler dileriz.")
        ])
        conn.commit()
        
    conn.close()

def bakiye_yeniden_hesapla(cari_id):
    """
    Belirli bir carinin tüm geçmiş işlemlerine göre bakiyesini baştan kusursuz hesaplar.
    - Satis_Faturasi: Müşteri bize borçlanır (+ bakiye artar)
    - Tahsilat: Müşteri öder (- bakiye düşer)
    - Alis_Faturasi: Tedarikçiye borçlanırız (- bakiye düşer / tedarikçiye borcumuz artar)
    - Odeme: Tedarikçiye öderiz (+ borcumuz kapanır / bakiye nötrlenir)
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT tur, tutar FROM islemler 
        WHERE cari_id = ? AND durum = 'Aktif'
    """, (cari_id,))
    
    rows = cursor.fetchall()
    net_bakiye = 0.0
    for r in rows:
        tur = r["tur"]
        tutar = float(r["tutar"] or 0.0)
        if tur == 'Satis_Faturasi':
            net_bakiye += tutar
        elif tur == 'Tahsilat':
            net_bakiye -= tutar
        elif tur == 'Alis_Faturasi':
            net_bakiye -= tutar
        elif tur == 'Odeme':
            net_bakiye += tutar
            
    cursor.execute("UPDATE cariler SET bakiye = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (net_bakiye, cari_id))
    conn.commit()
    conn.close()
    return net_bakiye

if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')
    init_db()
    print("Veritabani basariyla hazirlandi.")
