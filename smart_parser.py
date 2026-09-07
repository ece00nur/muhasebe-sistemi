import re
from datetime import datetime
from database import get_connection

DORSE_TIPLERI = [
    "Damper", "Lowbed", "Tenteli", "Konteyner Tasiyici", "Silobas",
    "Kapakli", "Frigo", "Platform", "Yedek Parca/Servis", "Diger"
]

GENEL_KELIMELER = {
    "ltd", "şti", "sti", "a.ş", "as", "san", "tic", "ve", "treyler", "dorse",
    "lojistik", "nakliyat", "sanayi", "ticaret", "otomotiv", "servis", "holding",
    "grup", "insaat", "inşaat", "tasimacilik", "taşımacılık", "limited", "sirketi", "şirketi"
}

def parse_tutar(text):
    """
    Metin içerisindeki tutarı akıllıca yakalar.
    Plaka ve şasi numaralarındaki sayıları tutar olarak algılamaz.
    TL / ₺ veya 'bin' ile belirtilen sayılara en yüksek önceliği verir.
    """
    # 1. Önce plakayı ve şasiyi metinden geçici olarak temizle ki sayıları tutarla karışmasın
    temiz_text = re.sub(r'\b\d{2}\s*[A-ZÇĞİÖŞÜ]{1,3}\s*\d{2,4}\b', ' ', text, flags=re.IGNORECASE)
    temiz_text = re.sub(r'(?:sasi|şasi|şase|vin)(?:\s*no)?[:\s]*[A-Z0-9]+', ' ', temiz_text, flags=re.IGNORECASE)
    
    # 2. '150 bin' veya '150k' veya '150 bin tl' yakalama
    bin_match = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:bin|k)\b', temiz_text, re.IGNORECASE)
    if bin_match:
        sayi_str = bin_match.group(1).replace(',', '.')
        try:
            return float(sayi_str) * 1000.0
        except ValueError:
            pass
            
    # 3. '1.5 milyon' yakalama
    milyon_match = re.search(r'(\d+(?:[.,]\d+)?)\s*milyon\b', temiz_text, re.IGNORECASE)
    if milyon_match:
        sayi_str = milyon_match.group(1).replace(',', '.')
        try:
            return float(sayi_str) * 1000000.0
        except ValueError:
            pass

    # 4. Açıkça 'TL', 'tl', '₺', 'lira' yanında geçen sayılar (En güvenilir)
    tl_match = re.search(r'(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?|\d+(?:[.,]\d{1,2})?)\s*(?:TL|tl|₺|lira)\b', temiz_text)
    if tl_match:
        return _temizle_ve_float_yap(tl_match.group(1))

    # 5. Genel sayı arama: 150000 veya 150.000 veya 150.000,50
    matches = re.findall(r'\b(?:\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:[.,]\d{1,2})?)\b', temiz_text)
    if matches:
        bulunanlar = []
        for m in matches:
            val = _temizle_ve_float_yap(m)
            if val > 0:
                bulunanlar.append(val)
        if bulunanlar:
            return max(bulunanlar)
            
    return 0.0

def _temizle_ve_float_yap(temiz):
    temiz = temiz.strip()
    if '.' in temiz and ',' in temiz:
        if temiz.rfind(',') > temiz.rfind('.'):
            temiz = temiz.replace('.', '').replace(',', '.')
        else:
            temiz = temiz.replace(',', '')
    elif '.' in temiz:
        parcalar = temiz.split('.')
        if len(parcalar[-1]) == 3:
            temiz = temiz.replace('.', '')
    elif ',' in temiz:
        parcalar = temiz.split(',')
        if len(parcalar[-1]) == 3:
            temiz = temiz.replace(',', '')
        else:
            temiz = temiz.replace(',', '.')
    try:
        return float(temiz)
    except ValueError:
        return 0.0

def cari_bul_veya_oner(text):
    """
    Veritabanındaki carileri tarayarak en yakın cariyi bulur.
    Genel kelimeleri ('treyler', 'nakliyat', 'ltd') yok sayar, asıl özel ada bakar.
    """
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, unvan, tip, telefon, bakiye FROM cariler")
    cariler = cursor.fetchall()
    conn.close()
    
    text_lower = text.lower()
    
    en_iyi_eslesme = None
    en_yuksek_puan = 0
    
    for c in cariler:
        unvan = c["unvan"].lower()
        # Tam unvan geçiyorsa en yüksek öncelik
        if unvan in text_lower:
            return {
                "id": c["id"],
                "unvan": c["unvan"],
                "tip": c["tip"],
                "bakiye": c["bakiye"],
                "mevcut": True
            }
            
        # Anlamlı özel kelimeleri ayıkla
        kelimeler = [k.strip() for k in re.split(r'[\s.,;]+', unvan) if len(k) > 2]
        ozel_kelimeler = [k for k in kelimeler if k not in GENEL_KELIMELER]
        
        puan = 0
        for k in ozel_kelimeler:
            if k in text_lower:
                puan += len(k) * 2 # Özel kelimelere yüksek ağırlık
                
        if puan > en_yuksek_puan and puan >= 6: # En az 3 harfli 1-2 özel kelime eşleşmeli
            en_yuksek_puan = puan
            en_iyi_eslesme = c
            
    if en_iyi_eslesme:
        return {
            "id": en_iyi_eslesme["id"],
            "unvan": en_iyi_eslesme["unvan"],
            "tip": en_iyi_eslesme["tip"],
            "bakiye": en_iyi_eslesme["bakiye"],
            "mevcut": True
        }
        
    # Eşleşme bulunamadıysa metinden firma adı tahmin et
    temizlenmis = re.sub(r'(?:sasi|şasi|şase|vin)(?:\s*no)?[:\s]*[A-Z0-9]+', ' ', text, flags=re.IGNORECASE)
    temizlenmis = re.sub(r'\b\d{2}\s*[A-ZÇĞİÖŞÜ]{1,3}\s*\d{2,4}\b', ' ', temizlenmis, flags=re.IGNORECASE)
    temizlenmis = re.sub(r'(\d+(?:[.,]\d+)?)\s*(?:bin|k|milyon|tl|₺)?', ' ', temizlenmis, flags=re.IGNORECASE)
    temizlenmis = re.sub(r'\b(fatura|kestik|kesildi|geldi|alis|alış|satis|satış|tahsilat|odeme|ödeme|havale|gonderdik|gönderdik|aldik|aldık|icin|için|nolu|damper|lowbed|treyler|dorse|sasi|şasi|plaka|ftr)\b', ' ', temizlenmis, flags=re.IGNORECASE)
    temizlenmis = re.sub(r'[:;,.!?-]', ' ', temizlenmis)
    temizlenmis = re.sub(r'\s+', ' ', temizlenmis).strip()
    
    if temizlenmis:
        onerilen_unvan = temizlenmis.title()
    else:
        onerilen_unvan = "Yeni Müşteri / Firma"
        
    return {
        "id": None,
        "unvan": onerilen_unvan,
        "tip": "Musteri",
        "bakiye": 0.0,
        "mevcut": False
    }

def dorse_ve_treyler_bilgisi_bul(text):
    """Metin içindeki dorse tipi, şasi no ve plakayı yakalar."""
    text_lower = text.lower()
    
    tespit_dorse = ""
    for d in DORSE_TIPLERI:
        d_lower = d.lower()
        if d_lower in text_lower:
            tespit_dorse = d
            break
            
    # Plaka yakalama (Örn: 34 ABC 123 veya 06 AA 9999 veya 42 KY 12)
    plaka_match = re.search(r'\b(\d{2}\s*[A-ZÇĞİÖŞÜ]{1,3}\s*\d{2,4})\b', text, re.IGNORECASE)
    plaka = plaka_match.group(1).upper() if plaka_match else ""
    
    # Şasi no yakalama (Şasi: TR... veya Şasi no ...)
    sasi_match = re.search(r'(?:sasi|şasi|şase|vin)(?:\s*no)?[:\s]*([A-Z0-9]{5,17})', text, re.IGNORECASE)
    sasi = sasi_match.group(1).upper() if sasi_match else ""
    
    return {
        "dorse_tipi": tespit_dorse,
        "plaka": plaka,
        "sasi_no": sasi
    }

def akilli_ayristir(metin):
    """
    Kullanıcının girdiği doğal metni anlar ve doğrudan işlenebilir bir işlem önerisi üretir.
    Örnekler:
    - 'Özdemir Lojistik 150000 fatura' -> Satis Faturasi
    - 'Kaya Lojistik 50 bin havale geldi' -> Tahsilat
    - 'Aksan Dingil 80.000 TL fatura geldi kampana için' -> Alis Faturasi
    - 'Demir Sac firmasına 35 bin ödeme yapıldı' -> Odeme
    """
    if not metin or not metin.strip():
        return None
        
    metin = metin.strip()
    metin_lower = metin.lower()
    
    # 1. Tutar bul
    tutar = parse_tutar(metin)
    
    # 2. Cari bul
    cari_bilgisi = cari_bul_veya_oner(metin)
    
    # 3. Treyler detayları bul
    dorse_detay = dorse_ve_treyler_bilgisi_bul(metin)
    
    # 4. İşlem türünü belirle
    # Öncelik kuralları:
    tur = "Satis_Faturasi" # Varsayılan: Müşteriye fatura kesilmesi
    tur_aciklamasi = "Satış Faturası (Bizim kestiğimiz fatura)"
    
    # Tahsilat anahtarları (Para Girişi)
    if any(k in metin_lower for k in ["tahsilat", "havale geldi", "para geldi", "hesaba gecti", "aldik", "bize odedi", "tahsil"]):
        tur = "Tahsilat"
        tur_aciklamasi = "Tahsilat / Para Girişi (Müşteri bize ödedi)"
    # Ödeme anahtarları (Para Çıkışı)
    elif any(k in metin_lower for k in ["odeme yaptik", "ödeme yaptık", "havale gonderdik", "havale gönderdik", "parayi gonderdik", "odedik", "çıkış"]):
        tur = "Odeme"
        tur_aciklamasi = "Ödeme / Para Çıkışı (Biz ödedik)"
    # Alış Faturası anahtarları (Bize kesilen fatura / Malzeme / Masraf)
    elif any(k in metin_lower for k in ["fatura geldi", "bize kesil", "alis", "alış", "malzeme faturasi", "sac faturasi", "dingil faturasi"]):
        tur = "Alis_Faturasi"
        tur_aciklamasi = "Alış Faturası (Tedarikçinin bize kestiği fatura)"
    # Satış Faturası anahtarları
    elif any(k in metin_lower for k in ["fatura kes", "kestik", "fatura", "satis", "satış", "ftr"]):
        tur = "Satis_Faturasi"
        tur_aciklamasi = "Satış Faturası (Bizim kestiğimiz fatura)"
    elif cari_bilgisi["tip"] == "Tedarikci":
        # Eğer tedarikçi seçilmişse ve ödeme denmemişse genelde alış faturasıdır
        tur = "Alis_Faturasi"
        tur_aciklamasi = "Alış Faturası (Tedarikçiden gelen fatura)"
        
    bugun = datetime.now().strftime("%Y-%m-%d")
    
    # Otomatik anlaşılır açıklama üret
    aciklama = metin
    
    return {
        "orijinal_metin": metin,
        "cari_id": cari_bilgisi["id"],
        "cari_unvan": cari_bilgisi["unvan"],
        "cari_mevcut": cari_bilgisi["mevcut"],
        "cari_tip": cari_bilgisi["tip"],
        "tur": tur,
        "tur_aciklamasi": tur_aciklamasi,
        "tutar": tutar,
        "tarih": bugun,
        "vade_tarihi": bugun,
        "dorse_tipi": dorse_detay["dorse_tipi"] or "Diger",
        "plaka": dorse_detay["plaka"],
        "sasi_no": dorse_detay["sasi_no"],
        "aciklama": aciklama,
        "onay_mesaji": f"✅ {cari_bilgisi['unvan']} için {tutar:,.2f} ₺ {tur_aciklamasi} kaydı oluşturulacak."
    }

if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')
        
    testler = [
        "Özdemir Lojistik 150000 fatura damper dorse satışı",
        "Kaya Nakliyat 50 bin havale geldi",
        "Aksan Dingil 85.000 TL fatura geldi 34 ABC 123",
        "Çeliksan firmasına 40 bin ödeme yaptık",
        "Yeni Karadeniz Treyler 320.000 TL ftr şasi: TR98765432"
    ]
    for t in testler:
        sonuc = akilli_ayristir(t)
        print(f"Girdi: {t}")
        print(f" -> Bulunan Cari: {sonuc['cari_unvan']} (ID: {sonuc['cari_id']})")
        print(f" -> Tür: {sonuc['tur']} | Tutar: {sonuc['tutar']} TL | Dorse: {sonuc['dorse_tipi']} | Plaka: {sonuc['plaka']}")
        print("-" * 50)
