import os
import io
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from database import get_connection, bakiye_yeniden_hesapla

# Kolaysoft ve diğer e-fatura sütun eşleşmeleri için olası anahtar kelimeler
KOLON_ESLESMELERI = {
    "belge_no": ["fatura no", "belge no", "ettn", "evrak no", "fatura numarası", "fatura no."],
    "tarih": ["fatura tarihi", "tarih", "düzenleme tarihi", "islem tarihi", "belge tarihi"],
    "unvan": ["alıcı ünvanı", "alici unvani", "müşteri", "musteri unvani", "firma adı", "cari ünvanı", "satıcı ünvanı", "ünvan", "cari adi"],
    "vergi_no": ["vkn", "tckn", "vergi no", "tc kimlik no", "vergi / tc kimlik no", "vkn/tckn"],
    "tutar": ["genel toplam", "ödenecek tutar", "toplam tutar", "net tutar", "tutar", "fatura tutarı", "toplam (tl)", "tutar (tl)"],
    "kdv": ["kdv tutarı", "kdv", "hesaplanan kdv"],
    "aciklama": ["açıklama", "aciklama", "not", "notlar", "mal/hizmet", "ürün"]
}

def _kolon_bul(header_row):
    """Excel başlık satırındaki hücreleri analiz ederek kolon indekslerini tespit eder."""
    indeksler = {}
    for col_idx, cell_value in enumerate(header_row):
        if cell_value is None:
            continue
        val = str(cell_value).strip().lower()
        val = val.replace("ı", "i").replace("ğ", "g").replace("ü", "u").replace("ş", "s").replace("ö", "o").replace("ç", "c")
        
        for key, possible_names in KOLON_ESLESMELERI.items():
            if key not in indeksler:
                for p in possible_names:
                    p_clean = p.replace("ı", "i").replace("ğ", "g").replace("ü", "u").replace("ş", "s").replace("ö", "o").replace("ç", "c")
                    if p_clean in val:
                        indeksler[key] = col_idx
                        break
    return indeksler

def kolaysoft_excel_ice_aktar(dosya_icerigi, dosya_adi=""):
    """
    KolaySoft'tan indirilen Gelen veya Giden e-fatura Excel dosyasını otomatik ayrıştırır.
    Carileri oluşturur/eşleştirir ve faturaları kaydeder.
    """
    wb = openpyxl.load_workbook(filename=io.BytesIO(dosya_icerigi), data_only=True)
    ws = wb.active
    
    # 1. Başlık satırını bul (Genelde 1. ile 5. satır arasındadır)
    header_row_idx = None
    kolon_haritasi = {}
    
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx > 10:
            break
        bulunan = _kolon_bul(row)
        if "unvan" in bulunan and "tutar" in bulunan:
            header_row_idx = r_idx
            kolon_haritasi = bulunan
            break
            
    if not header_row_idx:
        return {
            "basarili": False,
            "mesaj": "Excel dosyası içerisinde 'Ünvan/Firma Adı' ve 'Tutar' sütunları tespit edilemedi. Lütfen KolaySoft e-fatura listesi olduğundan emin olun."
        }
        
    conn = get_connection()
    cursor = conn.cursor()
    
    # Dosya adından veya içeriğinden Gelen / Giden Fatura tespiti
    dosya_adi_lower = dosya_adi.lower()
    varsayilan_tur = "Alis_Faturasi" if "gelen" in dosya_adi_lower else "Satis_Faturasi"
    
    toplam_satir = 0
    yeni_cari_sayisi = 0
    eklenen_fatura_sayisi = 0
    atlanan_mukerrer = 0
    etkilenen_cari_idler = set()
    
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx <= header_row_idx:
            continue
            
        unvan_val = row[kolon_haritasi["unvan"]] if "unvan" in kolon_haritasi and kolon_haritasi["unvan"] < len(row) else None
        tutar_val = row[kolon_haritasi["tutar"]] if "tutar" in kolon_haritasi and kolon_haritasi["tutar"] < len(row) else None
        
        if not unvan_val or not str(unvan_val).strip():
            continue
            
        unvan = str(unvan_val).strip()
        
        # Tutar çevirimi
        try:
            if isinstance(tutar_val, (int, float)):
                tutar = float(tutar_val)
            else:
                t_str = str(tutar_val).strip().replace(" ", "").replace("₺", "").replace("TL", "")
                if "," in t_str and "." in t_str:
                    if t_str.rfind(",") > t_str.rfind("."):
                        t_str = t_str.replace(".", "").replace(",", ".")
                    else:
                        t_str = t_str.replace(",", "")
                elif "," in t_str:
                    t_str = t_str.replace(",", ".")
                tutar = float(t_str)
        except Exception:
            continue
            
        if tutar <= 0:
            continue
            
        toplam_satir += 1
        
        # Ek kolonlar
        belge_no = ""
        if "belge_no" in kolon_haritasi and kolon_haritasi["belge_no"] < len(row):
            b_val = row[kolon_haritasi["belge_no"]]
            belge_no = str(b_val).strip() if b_val is not None else ""
            
        tarih = datetime.now().strftime("%Y-%m-%d")
        if "tarih" in kolon_haritasi and kolon_haritasi["tarih"] < len(row):
            t_val = row[kolon_haritasi["tarih"]]
            if isinstance(t_val, datetime):
                tarih = t_val.strftime("%Y-%m-%d")
            elif t_val:
                t_str = str(t_val).strip()
                # 07.09.2026 veya 2026-09-07 formatları
                m_tarih = re.search(r'(\d{1,4})[./-](\d{1,2})[./-](\d{2,4})', t_str)
                if m_tarih:
                    p1, p2, p3 = m_tarih.groups()
                    if len(p1) == 4:
                        tarih = f"{p1}-{p2.zfill(2)}-{p3.zfill(2)}"
                    else:
                        tarih = f"{p3}-{p2.zfill(2)}-{p1.zfill(2)}"
                        
        vkn = ""
        if "vergi_no" in kolon_haritasi and kolon_haritasi["vergi_no"] < len(row):
            v_val = row[kolon_haritasi["vergi_no"]]
            vkn = str(v_val).strip() if v_val is not None else ""
            
        aciklama = "KolaySoft İçe Aktarma"
        if "aciklama" in kolon_haritasi and kolon_haritasi["aciklama"] < len(row):
            a_val = row[kolon_haritasi["aciklama"]]
            if a_val:
                aciklama = str(a_val).strip()
                
        # 1. Cari eşleştirme veya oluşturma
        cari_id = None
        if vkn:
            cursor.execute("SELECT id FROM cariler WHERE vergi_no = ?", (vkn,))
            res = cursor.fetchone()
            if res:
                cari_id = res[0]
                
        if not cari_id:
            cursor.execute("SELECT id FROM cariler WHERE LOWER(unvan) = LOWER(?)", (unvan,))
            res = cursor.fetchone()
            if res:
                cari_id = res[0]
                
        if not cari_id:
            # Yeni cari kartı aç
            cari_tip = "Tedarikci" if varsayilan_tur == "Alis_Faturasi" else "Musteri"
            cursor.execute("""
                INSERT INTO cariler (unvan, tip, vergi_no, notlar)
                VALUES (?, ?, ?, 'KolaySoft aktarımında otomatik oluşturuldu')
            """, (unvan, cari_tip, vkn))
            cari_id = cursor.lastrowid
            yeni_cari_sayisi += 1
            
        etkilenen_cari_idler.add(cari_id)
        
        # 2. Mükerrer fatura kontrolü
        if belge_no:
            cursor.execute("SELECT id FROM islemler WHERE belge_no = ? AND cari_id = ?", (belge_no, cari_id))
            if cursor.fetchone():
                atlanan_mukerrer += 1
                continue
                
        # 3. Faturayı kaydet
        cursor.execute("""
            INSERT INTO islemler (cari_id, tur, belge_no, tarih, vade_tarihi, tutar, kdv_orani, aciklama, kolaysoft_id)
            VALUES (?, ?, ?, ?, ?, ?, 20.0, ?, ?)
        """, (cari_id, varsayilan_tur, belge_no, tarih, tarih, tutar, aciklama, belge_no))
        eklenen_fatura_sayisi += 1
        
    conn.commit()
    conn.close()
    
    # Etkilenen carilerin bakiyesini yeniden hesapla
    for cid in etkilenen_cari_idler:
        bakiye_yeniden_hesapla(cid)
        
    return {
        "basarili": True,
        "toplam_okunan": toplam_satir,
        "yeni_cari": yeni_cari_sayisi,
        "eklenen_fatura": eklenen_fatura_sayisi,
        "atlanan_mukerrer": atlanan_mukerrer,
        "mesaj": f"KolaySoft aktarımı tamamlandı! {eklenen_fatura_sayisi} fatura sisteme işlendi. ({yeni_cari_sayisi} yeni cari açıldı, {atlanan_mukerrer} mükerrer fatura atlandı)."
    }

def export_cariler_excel():
    """Tüm carileri ve güncel bakiyelerini profesyonelce biçimlendirilmiş bir Excel dosyası olarak üretir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, unvan, tip, yetkili, telefon, vergi_no, vergi_dairesi, bakiye, notlar
        FROM cariler ORDER BY unvan ASC
    """)
    cariler = cursor.fetchall()
    conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cari Hesap Listesi"
    ws.views.sheetView[0].showGridLines = True
    
    # Renk Stilleri
    baslik_font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    baslik_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Lacivert
    
    kolon_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    kolon_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid") # Canlı Mavi
    
    veri_font = Font(name="Segoe UI", size=10)
    kalin_font = Font(name="Segoe UI", size=10, bold=True)
    
    yesil_font = Font(name="Segoe UI", size=10, bold=True, color="15803D")
    kirmizi_font = Font(name="Segoe UI", size=10, bold=True, color="B91C1C")
    gri_font = Font(name="Segoe UI", size=10, color="6B7280")
    
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_ince = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    
    # 1. Başlık Banner
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = "TREYLER & DORSE SANAYİ - CARİ HESAP BAKİYE RAPORU"
    title_cell.font = baslik_font
    title_cell.fill = baslik_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40
    
    ws["A2"].value = f"Rapor Tarihi: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws["A2"].font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
    ws.row_dimensions[2].height = 20
    
    # 2. Sütun Başlıkları
    headers = [
        "Sıra", "Cari Ünvanı", "Cari Tipi", "Yetkili Kişi", 
        "Telefon", "Vergi No", "Güncel Bakiye (₺)", "Durum"
    ]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = kolon_font
        cell.fill = kolon_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_ince
    ws.row_dimensions[3].height = 28
    
    # 3. Veri Satırları
    row_idx = 4
    toplam_alacak = 0.0 # Müşterilerden alacağımız (+ bakiye)
    toplam_borc = 0.0   # Tedarikçilere borcumuz (- bakiye)
    
    for idx, c in enumerate(cariler, start=1):
        bakiye = float(c["bakiye"] or 0.0)
        
        durum_text = "Kapalı / 0.00 ₺"
        bakiye_font = gri_font
        if bakiye > 0.01:
            durum_text = "Alacağımız Var"
            bakiye_font = yesil_font
            toplam_alacak += bakiye
        elif bakiye < -0.01:
            durum_text = "Borcumuz Var"
            bakiye_font = kirmizi_font
            toplam_borc += abs(bakiye)
            
        tip_text = "Müşteri" if c["tip"] == "Musteri" else ("Tedarikçi" if c["tip"] == "Tedarikci" else "Diğer")
        
        row_data = [
            idx,
            c["unvan"],
            tip_text,
            c["yetkili"] or "-",
            c["telefon"] or "-",
            c["vergi_no"] or "-",
            bakiye,
            durum_text
        ]
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = veri_font
            cell.border = border_ince
            
            if row_idx % 2 == 0:
                cell.fill = zebra_fill
                
            if col_idx in [1, 3, 5, 6, 8]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 7: # Tutar
                cell.font = bakiye_font
                cell.number_format = '#,##0.00 "₺"'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
        ws.row_dimensions[row_idx].height = 22
        row_idx += 1
        
    # 4. Toplam Satırları
    ws.row_dimensions[row_idx].height = 26
    cell_label = ws.cell(row=row_idx, column=6, value="TOPLAM ALACAĞIMIZ:")
    cell_label.font = kalin_font
    cell_label.alignment = Alignment(horizontal="right", vertical="center")
    
    cell_val = ws.cell(row=row_idx, column=7, value=toplam_alacak)
    cell_val.font = yesil_font
    cell_val.number_format = '#,##0.00 "₺"'
    cell_val.alignment = Alignment(horizontal="right", vertical="center")
    
    row_idx += 1
    ws.row_dimensions[row_idx].height = 26
    cell_label2 = ws.cell(row=row_idx, column=6, value="TOPLAM BORCUMUZ:")
    cell_label2.font = kalin_font
    cell_label2.alignment = Alignment(horizontal="right", vertical="center")
    
    cell_val2 = ws.cell(row=row_idx, column=7, value=-toplam_borc)
    cell_val2.font = kirmizi_font
    cell_val2.number_format = '#,##0.00 "₺"'
    cell_val2.alignment = Alignment(horizontal="right", vertical="center")
    
    # Otomatik Sütun Genişlikleri
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in [1, 2]:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["H"].width = 18
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def export_cari_ekstre_excel(cari_id):
    """Seçilen carinin detaylı hesap ekstresini yürüyen bakiye ile birlikte Excel olarak üretir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM cariler WHERE id = ?", (cari_id,))
    cari = cursor.fetchone()
    
    if not cari:
        conn.close()
        return None
        
    cursor.execute("""
        SELECT * FROM islemler 
        WHERE cari_id = ? AND durum = 'Aktif'
        ORDER BY tarih ASC, id ASC
    """, (cari_id,))
    islemler = cursor.fetchall()
    conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hesap Ekstresi"
    ws.views.sheetView[0].showGridLines = True
    
    baslik_font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    baslik_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    
    kolon_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    kolon_fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    
    veri_font = Font(name="Segoe UI", size=10)
    kalin_font = Font(name="Segoe UI", size=10, bold=True)
    yesil_font = Font(name="Segoe UI", size=10, color="15803D", bold=True)
    kirmizi_font = Font(name="Segoe UI", size=10, color="B91C1C", bold=True)
    
    border_ince = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    
    # 1. Başlık
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = f"CARİ HESAP EKSTRESİ: {cari['unvan']}"
    title_cell.font = baslik_font
    title_cell.fill = baslik_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36
    
    ws["A2"].value = f"Yetkili: {cari['yetkili'] or '-'} | Tel: {cari['telefon'] or '-'} | Vergi No: {cari['vergi_no'] or '-'}"
    ws["A2"].font = Font(name="Segoe UI", size=10, color="475569")
    ws.row_dimensions[2].height = 20
    
    ws["A3"].value = f"Ekstre Çıkış Tarihi: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws["A3"].font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
    ws.row_dimensions[3].height = 18
    
    # 2. Tablo Başlıkları
    headers = [
        "Tarih", "İşlem Türü", "Belge No", "Treyler / Dorse", 
        "Şasi No / Plaka", "Açıklama", "Borç / Giriş (₺)", "Alacak / Çıkış (₺)", "Kalan Bakiye (₺)"
    ]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=col_idx, value=h)
        cell.font = kolon_font
        cell.fill = kolon_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_ince
    ws.row_dimensions[5].height = 26
    
    # 3. Yürüyen Bakiye Hesaplama & Satırlar
    row_idx = 6
    yuruyen_bakiye = 0.0
    toplam_borc = 0.0
    toplam_alacak = 0.0
    
    tur_isimleri = {
        "Satis_Faturasi": "Satış Faturası",
        "Tahsilat": "Tahsilat (Havale/Nakit)",
        "Alis_Faturasi": "Alış Faturası",
        "Odeme": "Ödeme Yapıldı"
    }
    
    for islem in islemler:
        tur = islem["tur"]
        tutar = float(islem["tutar"] or 0.0)
        
        borc_tutar = 0.0
        alacak_tutar = 0.0
        
        if tur == "Satis_Faturasi":
            borc_tutar = tutar
            yuruyen_bakiye += tutar
            toplam_borc += tutar
        elif tur == "Tahsilat":
            alacak_tutar = tutar
            yuruyen_bakiye -= tutar
            toplam_alacak += tutar
        elif tur == "Alis_Faturasi":
            alacak_tutar = tutar
            yuruyen_bakiye -= tutar
            toplam_alacak += tutar
        elif tur == "Odeme":
            borc_tutar = tutar
            yuruyen_bakiye += tutar
            toplam_borc += tutar
            
        sasi_plaka = f"{islem['sasi_no'] or ''} {islem['plaka'] or ''}".strip() or "-"
        
        row_data = [
            islem["tarih"],
            tur_isimleri.get(tur, tur),
            islem["belge_no"] or "-",
            islem["dorse_tipi"] or "-",
            sasi_plaka,
            islem["aciklama"] or "-",
            borc_tutar if borc_tutar > 0 else "",
            alacak_tutar if alacak_tutar > 0 else "",
            yuruyen_bakiye
        ]
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = veri_font
            cell.border = border_ince
            
            if col_idx in [1, 2, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [7, 8]:
                cell.number_format = '#,##0.00 "₺"'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx == 9:
                cell.number_format = '#,##0.00 "₺"'
                cell.font = yesil_font if yuruyen_bakiye >= 0 else kirmizi_font
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
        ws.row_dimensions[row_idx].height = 22
        row_idx += 1
        
    # Toplam Satırı
    ws.row_dimensions[row_idx].height = 26
    ws.cell(row=row_idx, column=6, value="TOPLAMLAR:").font = kalin_font
    ws.cell(row=row_idx, column=6).alignment = Alignment(horizontal="right", vertical="center")
    
    c_tb = ws.cell(row=row_idx, column=7, value=toplam_borc)
    c_tb.font = kalin_font
    c_tb.number_format = '#,##0.00 "₺"'
    c_tb.alignment = Alignment(horizontal="right", vertical="center")
    
    c_ta = ws.cell(row=row_idx, column=8, value=toplam_alacak)
    c_ta.font = kalin_font
    c_ta.number_format = '#,##0.00 "₺"'
    c_ta.alignment = Alignment(horizontal="right", vertical="center")
    
    c_net = ws.cell(row=row_idx, column=9, value=yuruyen_bakiye)
    c_net.font = yesil_font if yuruyen_bakiye >= 0 else kirmizi_font
    c_net.number_format = '#,##0.00 "₺"'
    c_net.alignment = Alignment(horizontal="right", vertical="center")
    
    # Kolon Genişlikleri
    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 36
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 18
    ws.column_dimensions["I"].width = 20
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')
    res = export_cariler_excel()
    print(f"Excel aktarma basarili. Olusan boyut: {len(res.getvalue())} bytes")
