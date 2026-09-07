import os
import shutil
from datetime import datetime
from typing import Optional, List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from database import init_db, get_connection, bakiye_yeniden_hesapla, DB_PATH
from smart_parser import akilli_ayristir
from excel_service import kolaysoft_excel_ice_aktar, export_cariler_excel, export_cari_ekstre_excel

# Veritabanını hazırla
init_db()

app = FastAPI(title="Treyler Cari & Fatura Takip Sistemi", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Statik ve şablon dizinleri
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# Pydantic Şemaları
class CariCreate(BaseModel):
    unvan: str
    tip: Optional[str] = "Musteri"
    yetkili: Optional[str] = ""
    telefon: Optional[str] = ""
    eposta: Optional[str] = ""
    vergi_no: Optional[str] = ""
    vergi_dairesi: Optional[str] = ""
    adres: Optional[str] = ""
    notlar: Optional[str] = ""

class IslemCreate(BaseModel):
    cari_id: Optional[int] = None
    cari_unvan: Optional[str] = None
    cari_tip: Optional[str] = "Musteri"
    tur: str # Satis_Faturasi, Alis_Faturasi, Tahsilat, Odeme
    belge_no: Optional[str] = ""
    tarih: Optional[str] = ""
    vade_tarihi: Optional[str] = ""
    tutar: float
    kdv_orani: Optional[float] = 20.0
    tevkifat: Optional[str] = "Yok"
    odeme_yontemi: Optional[str] = "Havale/EFT"
    sasi_no: Optional[str] = ""
    dorse_tipi: Optional[str] = "Diger"
    plaka: Optional[str] = ""
    aciklama: Optional[str] = ""

class SmartParseRequest(BaseModel):
    text: str

# ----------------- SAYFA ENDPOINT -----------------
@app.get("/", response_class=HTMLResponse)
def index_page():
    index_file = os.path.join(TEMPLATES_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Sistem yükleniyor, lütfen index.html dosyasını kontrol edin.</h1>"

# ----------------- İSTATİSTİKLER -----------------
@app.get("/api/istatistikler")
def get_istatistikler():
    conn = get_connection()
    cursor = conn.cursor()
    
    # Tüm carilerin bakiyeleri
    cursor.execute("SELECT bakiye FROM cariler")
    cariler = cursor.fetchall()
    
    toplam_alacak = sum(r["bakiye"] for r in cariler if r["bakiye"] > 0)
    toplam_borc = sum(abs(r["bakiye"]) for r in cariler if r["bakiye"] < 0)
    
    # Bu ayki satış ve tahsilatlar
    bu_ay = datetime.now().strftime("%Y-%m")
    cursor.execute("""
        SELECT tur, SUM(tutar) as toplam 
        FROM islemler 
        WHERE tarih LIKE ? AND durum = 'Aktif'
        GROUP BY tur
    """, (f"{bu_ay}%",))
    aylik_islemler = {r["tur"]: r["toplam"] for r in cursor.fetchall()}
    
    conn.close()
    
    return {
        "toplam_alacak": toplam_alacak, # Müşterilerden alacağımız para
        "toplam_borc": toplam_borc,     # Tedarikçilere borcumuz
        "net_durum": toplam_alacak - toplam_borc,
        "bu_ay_satis": aylik_islemler.get("Satis_Faturasi", 0.0),
        "bu_ay_tahsilat": aylik_islemler.get("Tahsilat", 0.0),
        "bu_ay_alis": aylik_islemler.get("Alis_Faturasi", 0.0),
        "bu_ay_odeme": aylik_islemler.get("Odeme", 0.0)
    }

# ----------------- CARİLER ENDPOINTS -----------------
@app.get("/api/cariler")
def get_cariler(q: Optional[str] = None, tip: Optional[str] = None):
    conn = get_connection()
    cursor = conn.cursor()
    
    query = "SELECT * FROM cariler WHERE 1=1"
    params = []
    
    if q and q.strip():
        q_clean = f"%{q.strip()}%"
        query += " AND (unvan LIKE ? OR yetkili LIKE ? OR telefon LIKE ? OR vergi_no LIKE ?)"
        params.extend([q_clean, q_clean, q_clean, q_clean])
        
    if tip and tip != "Tumu":
        query += " AND tip = ?"
        params.append(tip)
        
    query += " ORDER BY ABS(bakiye) DESC, unvan ASC"
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows

@app.post("/api/cariler")
def create_cari(data: CariCreate):
    conn = get_connection()
    cursor = conn.cursor()
    
    unvan = data.unvan.strip()
    if not unvan:
        raise HTTPException(status_code=400, detail="Firma / Cari ünvanı zorunludur.")
        
    cursor.execute("""
        INSERT INTO cariler (unvan, tip, yetkili, telefon, eposta, vergi_no, vergi_dairesi, adres, notlar)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (unvan, data.tip, data.yetkili, data.telefon, data.eposta, data.vergi_no, data.vergi_dairesi, data.adres, data.notlar))
    
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return {"id": new_id, "mesaj": f"'{unvan}' başarıyla eklendi."}

@app.put("/api/cariler/{cari_id}")
def update_cari(cari_id: int, data: CariCreate):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE cariler SET 
            unvan = ?, tip = ?, yetkili = ?, telefon = ?, eposta = ?, 
            vergi_no = ?, vergi_dairesi = ?, adres = ?, notlar = ?, updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
    """, (data.unvan, data.tip, data.yetkili, data.telefon, data.eposta, data.vergi_no, data.vergi_dairesi, data.adres, data.notlar, cari_id))
    conn.commit()
    conn.close()
    return {"mesaj": "Cari bilgileri güncellendi."}

@app.delete("/api/cariler/{cari_id}")
def delete_cari(cari_id: int):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM islemler WHERE cari_id = ?", (cari_id,))
    cursor.execute("DELETE FROM cariler WHERE id = ?", (cari_id,))
    conn.commit()
    conn.close()
    return {"mesaj": "Cari ve bağlı tüm geçmiş hareketleri silindi."}

@app.get("/api/cariler/{cari_id}/ekstre")
def get_cari_ekstre(cari_id: int):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM cariler WHERE id = ?", (cari_id,))
    cari = cursor.fetchone()
    if not cari:
        conn.close()
        raise HTTPException(status_code=404, detail="Cari bulunamadı.")
        
    cursor.execute("""
        SELECT * FROM islemler 
        WHERE cari_id = ? AND durum = 'Aktif'
        ORDER BY tarih ASC, id ASC
    """, (cari_id,))
    islemler = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    # Yürüyen bakiye hesapla
    yuruyen_bakiye = 0.0
    for islem in islemler:
        tur = islem["tur"]
        tutar = float(islem["tutar"] or 0.0)
        if tur in ["Satis_Faturasi", "Odeme"]:
            yuruyen_bakiye += tutar
        elif tur in ["Tahsilat", "Alis_Faturasi"]:
            yuruyen_bakiye -= tutar
        islem["yuruyen_bakiye"] = yuruyen_bakiye
        
    return {
        "cari": dict(cari),
        "islemler": islemler,
        "net_bakiye": yuruyen_bakiye
    }

# ----------------- İŞLEMLER ENDPOINTS -----------------
@app.get("/api/islemler")
def get_islemler(limit: int = 50, cari_id: Optional[int] = None):
    conn = get_connection()
    cursor = conn.cursor()
    
    query = """
        SELECT i.*, c.unvan as cari_unvan, c.tip as cari_tip
        FROM islemler i
        JOIN cariler c ON i.cari_id = c.id
        WHERE i.durum = 'Aktif'
    """
    params = []
    if cari_id:
        query += " AND i.cari_id = ?"
        params.append(cari_id)
        
    query += " ORDER BY i.tarih DESC, i.id DESC LIMIT ?"
    params.append(limit)
    
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows

@app.post("/api/islemler")
def create_islem(data: IslemCreate):
    conn = get_connection()
    cursor = conn.cursor()
    
    cari_id = data.cari_id
    
    # Eğer cari_id yoksa ama firma adı verilmişse cariyi anında oluştur!
    if not cari_id and data.cari_unvan:
        unvan_temiz = data.cari_unvan.strip()
        cursor.execute("SELECT id FROM cariler WHERE LOWER(unvan) = LOWER(?)", (unvan_temiz,))
        mevcut = cursor.fetchone()
        if mevcut:
            cari_id = mevcut[0]
        else:
            cursor.execute("""
                INSERT INTO cariler (unvan, tip, notlar)
                VALUES (?, ?, 'Hızlı işlem sırasında otomatik oluşturuldu')
            """, (unvan_temiz, data.cari_tip or "Musteri"))
            cari_id = cursor.lastrowid
            
    if not cari_id:
        conn.close()
        raise HTTPException(status_code=400, detail="Lütfen bir cari/firma seçin veya yeni firma adını yazın.")
        
    tarih = data.tarih or datetime.now().strftime("%Y-%m-%d")
    vade = data.vade_tarihi or tarih
    
    cursor.execute("""
        INSERT INTO islemler (
            cari_id, tur, belge_no, tarih, vade_tarihi, tutar, kdv_orani, tevkifat,
            odeme_yontemi, sasi_no, dorse_tipi, plaka, aciklama
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        cari_id, data.tur, data.belge_no, tarih, vade, data.tutar,
        data.kdv_orani, data.tevkifat, data.odeme_yontemi,
        data.sasi_no, data.dorse_tipi, data.plaka, data.aciklama
    ))
    
    new_islem_id = cursor.lastrowid
    conn.commit()
    conn.close()
    
    # Bakiye güncellemesi
    yeni_bakiye = bakiye_yeniden_hesapla(cari_id)
    
    return {
        "id": new_islem_id,
        "cari_id": cari_id,
        "yeni_bakiye": yeni_bakiye,
        "mesaj": "İşlem başarıyla kaydedildi."
    }

@app.delete("/api/islemler/{islem_id}")
def delete_islem(islem_id: int):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT cari_id FROM islemler WHERE id = ?", (islem_id,))
    res = cursor.fetchone()
    if not res:
        conn.close()
        raise HTTPException(status_code=404, detail="İşlem bulunamadı.")
    cari_id = res[0]
    
    cursor.execute("DELETE FROM islemler WHERE id = ?", (islem_id,))
    conn.commit()
    conn.close()
    
    yeni_bakiye = bakiye_yeniden_hesapla(cari_id)
    return {"mesaj": "İşlem başarıyla silindi (Geri alındı).", "yeni_bakiye": yeni_bakiye}

# ----------------- AKILLI DOĞAL DİL PARSER -----------------
@app.post("/api/smart-parse")
def parse_smart_input(req: SmartParseRequest):
    sonuc = akilli_ayristir(req.text)
    if not sonuc:
        return {"basarili": False, "mesaj": "Anlaşılamadı. Lütfen firma adı ve tutar belirtin."}
    sonuc["basarili"] = True
    return sonuc

# ----------------- KOLAYSOFT VE EXCEL AKTARIMLARI -----------------
@app.post("/api/import/kolaysoft")
async def import_kolaysoft(dosya: UploadFile = File(...)):
    icerik = await dosya.read()
    sonuc = kolaysoft_excel_ice_aktar(icerik, dosya.filename)
    return sonuc

@app.get("/api/export/cariler")
def export_cariler():
    excel_stream = export_cariler_excel()
    dosya_adi = f"Cari_Listesi_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(
        excel_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={dosya_adi}"}
    )

@app.get("/api/export/ekstre/{cari_id}")
def export_ekstre(cari_id: int):
    excel_stream = export_cari_ekstre_excel(cari_id)
    if not excel_stream:
        raise HTTPException(status_code=404, detail="Cari bulunamadı.")
    dosya_adi = f"Cari_Ekstre_{cari_id}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(
        excel_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={dosya_adi}"}
    )

# ----------------- YEDEKLEME VE AYARLAR -----------------
@app.get("/api/backup")
def backup_database():
    """Veritabanının anlık kopyasını indirir."""
    if not os.path.exists(DB_PATH):
        raise HTTPException(status_code=404, detail="Veritabanı dosyası bulunamadı.")
    tarih_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    yedek_ad = f"Treyler_Muhasebe_Yedek_{tarih_str}.db"
    return FileResponse(DB_PATH, filename=yedek_ad, media_type="application/octet-stream")

@app.get("/api/ayarlar")
def get_ayarlar():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT anahtar, deger FROM ayarlar")
    rows = cursor.fetchall()
    conn.close()
    return {r["anahtar"]: r["deger"] for r in rows}

@app.post("/api/ayarlar")
def update_ayarlar(ayarlar: dict):
    conn = get_connection()
    cursor = conn.cursor()
    for k, v in ayarlar.items():
        cursor.execute("INSERT OR REPLACE INTO ayarlar (anahtar, deger) VALUES (?, ?)", (k, str(v)))
    conn.commit()
    conn.close()
    return {"mesaj": "Ayarlar kaydedildi."}

if __name__ == "__main__":
    import uvicorn
    import sys
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')
    port = int(os.environ.get("PORT", 8000))
    print(f"Treyler Cari & Fatura Sunucusu Başlatılıyor: http://127.0.0.1:{port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
