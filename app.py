import io
import re
from datetime import datetime
from flask import Flask, request, send_file, render_template_string
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

app = Flask(__name__)

# --- GÖRSEL ARAYÜZ (HTML/CSS) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Happy Center Rapor Sistemi</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f9; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .container { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center; width: 450px; }
        h1 { color: #e67e22; margin: 0; font-size: 32px; }
        h2 { color: #555; margin-top: 5px; font-size: 18px; font-weight: normal;}
        .info { background-color: #e8f8f5; color: #0e6251; padding: 10px; border-radius: 5px; font-size: 13px; margin-bottom: 20px; border: 1px solid #a2d9ce;}
        .upload-btn-wrapper { position: relative; overflow: hidden; display: inline-block; margin-bottom: 20px; width: 100%; }
        .btn { border: 2px dashed #e67e22; color: #e67e22; background-color: white; padding: 20px; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; transition: 0.3s; width: 100%; box-sizing: border-box;}
        .btn:hover { background-color: #fdf2e9; }
        input[type=file] { font-size: 100px; position: absolute; left: 0; top: 0; opacity: 0; cursor: pointer; height: 100%; width: 100%; }
        .file-name { margin-top: -10px; margin-bottom: 15px; color: #333; font-weight: bold; font-size: 14px; display: block; min-height: 20px;}
        .convert-btn { background-color: #27ae60; color: white; border: none; padding: 15px; width: 100%; border-radius: 8px; font-size: 18px; font-weight: bold; cursor: pointer; box-shadow: 0 4px #1e8449; transition: 0.1s; }
        .convert-btn:active { box-shadow: 0 2px #1e8449; transform: translateY(2px); }
        .convert-btn:hover { background-color: #2ecc71; }
    </style>
</head>
<body>
    <div class="container">
        <h1>HAPPY CENTER</h1>
        <h2>Şube Açılış/Kapanış Raporu</h2>
        
        <div class="info">
            Sisteme sadece HTML dosyasını yükleyin, Excel otomatik oluşacaktır.
        </div>

        <form action="/" method="post" enctype="multipart/form-data">
            <div class="upload-btn-wrapper">
                <div class="btn">📂 Dosya Seçmek İçin Tıklayın</div>
                <input type="file" name="file" required onchange="document.getElementById('fname').innerText = this.files[0].name; document.querySelector('.btn').style.borderColor='#27ae60'; document.querySelector('.btn').style.color='#27ae60';">
            </div>
            <span id="fname" class="file-name"></span>
            
            <button type="submit" class="convert-btn">Excel'e Dönüştür ve İndir</button>
        </form>
    </div>
</body>
</html>
"""

# --- YARDIMCI FONKSİYON: Türkçe Tarih Çevirici ---
def tarihi_formatla(tarih_metni):
    """ '26 Kasım 2025' formatındaki metni '26.11.2025' formatına çevirir """
    try:
        aylar = {
            "Ocak": "01", "Şubat": "02", "Mart": "03", "Nisan": "04", "Mayıs": "05", "Haziran": "06",
            "Temmuz": "07", "Ağustos": "08", "Eylül": "09", "Ekim": "10", "Kasım": "11", "Aralık": "12",
            "January": "01", "February": "02", "March": "03", "April": "04", "May": "05", "June": "06",
            "July": "07", "August": "08", "September": "09", "October": "10", "November": "11", "December": "12"
        }
        
        # Gereksiz boşluk ve iki nokta temizliği
        temiz_metin = tarih_metni.replace(":", "").strip()
        parcalar = temiz_metin.split() # ['26', 'Kasım', '2025']
        
        if len(parcalar) >= 3:
            gun = parcalar[0].zfill(2) # 2 -> 02 yapar
            ay_isim = parcalar[1]
            yil = parcalar[2]
            
            ay_no = aylar.get(ay_isim, "00")
            if ay_no != "00":
                return f"{gun}.{ay_no}.{yil}"
        
        return temiz_metin # Çeviremezse olduğu gibi dönsün
    except:
        return tarih_metni

# --- PYTHON MANTIĞI (EXCEL OLUŞTURMA) ---
def process_html_to_excel(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    satirlar = soup.find_all('tr')
    
    sube_verileri = {}
    aktif_sube = None
    rapor_tarihi_html = None # HTML'den çekilecek tarih

    # HTML Analizi
    for satir in satirlar:
        satir_metni = satir.get_text(" ", strip=True)
        hucreler = satir.find_all('td')

        # --- 0. TARİH BULMA ---
        if "Rapor Tarihi" in satir_metni:
            for td in hucreler:
                txt = td.get_text(strip=True)
                # 'Rapor Tarihi' yazısı olmayan ama dolu olan hücre tarihtir
                if txt and "Rapor Tarihi" not in txt:
                    rapor_tarihi_html = tarihi_formatla(txt) # 26.11.2025'e çevir
                    break
            continue

        # --- 1. ŞUBE BULMA ---
        if "Firma Ünvanı" in satir_metni:
            for td in hucreler:
                txt = td.get_text(strip=True)
                if txt and "Firma Ünvanı" not in txt and len(txt) > 2:
                    aktif_sube = txt
                    break
            continue
        
        if not "Firma Ünvanı" in satir_metni:
            genis_hucre = satir.find('td', attrs={'colspan': '28'})
            if genis_hucre:
                txt = genis_hucre.get_text(strip=True)
                if txt: aktif_sube = txt

        # --- 2. VERİ İŞLEME ---
        if aktif_sube and ("SİSTEM KAPATILDI" in satir_metni or "SİSTEM KURULDU" in satir_metni):
            if aktif_sube not in sube_verileri:
                sube_verileri[aktif_sube] = {"acilis_saat": "", "acilis_kisi": "", "kapanis_saat": "", "kapanis_kisi": ""}

            # Saat Bulma
            saat_match = re.search(r'\b([0-9]{1,2}):([0-9]{2})\b', satir_metni)
            saat = saat_match.group(0) if saat_match else ""

            # Personel Bulma ve Temizleme
            ham_veri = ""
            for td in hucreler:
                if "SİSTEM" in td.get_text():
                    ham_veri = td.get_text(strip=True)
                    break
            if not ham_veri: ham_veri = satir_metni
            
            personel = ""
            if "SİSTEM" in ham_veri:
                personel = ham_veri.split("SİSTEM")[0].strip(" .")
                # İsim başındaki sayıları temizle (002. vb)
                personel = re.sub(r'^\d+\.\s*', '', personel)

            # Durum Kontrolü
            if "SİSTEM KAPATILDI" in satir_metni: # Mağaza Açıldı
                sube_verileri[aktif_sube]["acilis_saat"] = saat
                sube_verileri[aktif_sube]["acilis_kisi"] = personel
            elif "SİSTEM KURULDU" in satir_metni: # Mağaza Kapandı
                sube_verileri[aktif_sube]["kapanis_saat"] = saat
                sube_verileri[aktif_sube]["kapanis_kisi"] = personel

    # --- EXCEL OLUŞTURMA ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Happy Center Rapor"

    # Stiller
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # SARI
    bold_font = Font(bold=True, name='Calibri', size=11)
    title_font = Font(bold=True, name='Calibri', size=14)
    center_align = Alignment(horizontal='center', vertical='center')

    # Başlık Alanı
    ws.merge_cells('B1:H1')
    ws['B1'] = "HAPPY CENTER ŞUBELERİN AÇILIŞ KAPANIŞ SAATLERİ"
    ws['B1'].font = title_font
    ws['B1'].alignment = center_align

    # TARİH BAŞLIĞI (Dinamik)
    # Eğer HTML'den tarih bulduysa onu kullan, bulamazsa bugünü kullan
    tarih_str = rapor_tarihi_html if rapor_tarihi_html else datetime.now().strftime("%d.%m.%Y")
    
    ws.merge_cells('B2:H2')
    ws['B2'] = f"{tarih_str} HAPPY CENTER MAĞAZA AÇILIŞ VE KAPANIŞLARI"
    ws['B2'].font = bold_font
    ws['B2'].alignment = center_align

    # Tablo Başlıkları
    headers_config = [
        ('A4:A5', 'SIRA NO'),
        ('B4:B5', 'ŞUBE ADI'),
        ('C4:D4', 'ŞUBEYİ AÇAN'), ('C5', 'SAAT'), ('D5', 'PERSONEL'),
        ('E4:F4', 'ŞUBEYİ KAPATAN'), ('E5', 'SAAT'), ('F5', 'PERSONEL'),
        ('G4:G5', 'AÇIKLAMA')
    ]
    
    for rng, val in headers_config:
        if ':' in rng: ws.merge_cells(rng)
        cell = ws[rng.split(':')[0]]
        cell.value = val
        cell.fill = header_fill
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = thin_border

    # Başlık Kenarlıkları Tamamlama
    for r in ws['A4:G5']:
        for c in r: c.border = thin_border

    # Verileri Yazdırma
    start_row = 6
    sira_no = 1
    sorted_subeler = sorted(sube_verileri.keys())

    for sube in sorted_subeler:
        data = sube_verileri[sube]
        
        # SIRA NO
        c = ws.cell(row=start_row, column=1, value=sira_no)
        c.border = thin_border
        c.alignment = center_align
        
        # ŞUBE ADI
        c = ws.cell(row=start_row, column=2, value=sube)
        c.font = bold_font
        c.border = thin_border

        # AÇILIŞ SAAT (Yeşil & Ortalı)
        c = ws.cell(row=start_row, column=3, value=data['acilis_saat'])
        c.font = Font(color="006100") 
        c.alignment = center_align
        c.border = thin_border
        
        # AÇILIŞ PERSONEL (Yeşil)
        c = ws.cell(row=start_row, column=4, value=data['acilis_kisi'])
        c.font = Font(color="006100")
        c.border = thin_border

        # KAPANIŞ SAAT (Lacivert & Ortalı)
        c = ws.cell(row=start_row, column=5, value=data['kapanis_saat'])
        c.font = Font(color="000080")
        c.alignment = center_align
        c.border = thin_border
        
        # KAPANIŞ PERSONEL (Lacivert)
        c = ws.cell(row=start_row, column=6, value=data['kapanis_kisi'])
        c.font = Font(color="000080")
        c.border = thin_border
        
        # AÇIKLAMA
        ws.cell(row=start_row, column=7, value="").border = thin_border

        start_row += 1
        sira_no += 1

    # Genişlik Ayarları
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 15

    # Belleğe kaydetme
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- FLASK BAŞLATMA ---
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            return "Lütfen bir dosya seçin!", 400
        
        try:
            html_content = file.read().decode('utf-8')
            excel_file = process_html_to_excel(html_content)
            
            return send_file(
                excel_file, 
                as_attachment=True, 
                # Dosya adını da dinamik yapabiliriz ama şimdilik "Rapor_Tarih" formatında kalsın
                download_name=f"Happy_Center_Rapor_{datetime.now().strftime('%d_%m')}.xlsx",
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            return f"Bir hata oluştu: {str(e)}", 500

    return render_template_string(HTML_TEMPLATE)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
