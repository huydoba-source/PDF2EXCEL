import gc
import streamlit as st
import pandas as pd
import os
import io
import re
import time
import requests  
import pdfplumber
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- THƯ VIỆN BỔ SUNG CHO OCR ---
import pytesseract
from pytesseract import Output
from PIL import Image, ImageDraw, ImageOps
from docling.document_converter import DocumentConverter

# [LƯU Ý]: Sửa lại đường dẫn Tesseract trên máy bạn nếu cần
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ==========================================
# 1. CẤU HÌNH CỘT DỮ LIỆU ĐẦU RA 
# ==========================================
COLUMNS = [
    "Form",
    "Reference No",
    "Original CO Reference Number",
    "Item Number",
    "English description",
    "Quantity",
    "UOM",
    "USD",
    "Origin criteria (see Overleaf Notes)",
    "IMPORTING COUNTRY HS CODE",
    "EXPORTING COUNTRY HS CODE",
    "Invoice Number",
    "Date of invoices",
    "CARTON",
    "Original CO Issuance Date",
    "Issuing Authority",
    "Date of certification",
    "Products consigned from (Exporter's business name, address, country)",
    "Products consigned to (Consignee's name, address, country)",
    "Means of transport and route (as far as known)",
    "Produced in",
    "Exported to",
    "Marks and numbers on packages",
    "Box 13"
]

DECATHLON_BLUE = "#0082C3"
DECATHLON_DARK = "#1F2937"
BG_LIGHT = "#F9FAFB"

# ==========================================
# HÀM GỬI EMAIL THÔNG BÁO (SMTP)
# ==========================================
def send_email_notification():
    SENDER_EMAIL = "dobahuy7@gmail.com" 
    SENDER_PASSWORD = "kwyv yjud qvhy ehiq" 
    RECEIVER_EMAIL = "huy.doba@decathlon.com" 
    
    if "nhap_gmail_ao" in SENDER_EMAIL: return

    try:
        msg = MIMEMultipart()
        msg['From'] = f"Hệ thống Form E/D <{SENDER_EMAIL}>"
        msg['To'] = RECEIVER_EMAIL
        msg['Subject'] = "🚨 Thông báo: Có người truy cập Web Form E/D"
        body = f"Chào Huy,\n\nVừa có một người dùng mới truy cập vào ứng dụng trích xuất PDF lúc {time.strftime('%Y-%m-%d %H:%M:%S')}."
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        print(f"[!] Lỗi khi gửi email SMTP: {e}")

# ==========================================
# CÁC HÀM TIỆN ÍCH CHUNG
# ==========================================
def format_to_dd_mm_yyyy(date_str):
    if not date_str: return ""
    date_str = date_str.strip()
    months_map = {
        "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
        "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",
        "january": "01", "february": "02", "march": "03", "april": "04", "june": "06",
        "july": "07", "august": "08", "september": "09", "october": "10", "november": "11", "december": "12"
    }
    match = re.search(r'(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})', date_str)
    if match: return f"{int(match.group(1)):02d}-{int(match.group(2)):02d}-{match.group(3)}"
    match = re.search(r'(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})', date_str)
    if match: return f"{int(match.group(3)):02d}-{int(match.group(2)):02d}-{match.group(1)}"
    match = re.search(r'(\d{1,2})\s*[\-\s]\s*([A-Za-z]+)\s*[\-\s]\s*(\d{4})', date_str)
    if match:
        m_str = match.group(2).lower()
        if m_str in months_map: return f"{int(match.group(1)):02d}-{months_map[m_str]}-{match.group(3)}"
    return date_str

def clean_text(text):
    if not text: return ""
    text = re.sub(r'(?i)\b(?:page\s*\d+\s*of\s*\d+|page\s*\d+\s*of|\d+\s*of\s*\d+|\d+\s*of|page\s*\d+|of\s*\d+)\b', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'^:\s*', '', text)
    return text.strip() if re.search(r'[A-Za-z0-9]', text) else ""

# ==========================================
# 2. XỬ LÝ ĐẶC BIỆT CHO FILE PDF LÀ ẢNH SCAN TOÀN BỘ
# ==========================================
def extract_box13_scanned(page):
    """ Dùng thuật toán Neo hình học để lấy Box 13 mà không cần lưu ảnh """
    try:
        bbox_13 = (0, page.height - 250, page.width, page.height)
        img = page.crop(bbox_13).to_image(resolution=300).original
        ocr_data = pytesseract.image_to_data(img, output_type=Output.DICT, config='--oem 3 --psm 11')
        
        words = []
        for i in range(len(ocr_data['text'])):
            text = ocr_data['text'][i].strip()
            if len(text) >= 2:
                words.append({'text': text, 'x0': ocr_data['left'][i], 'top': ocr_data['top'][i], 'x1': ocr_data['left'][i] + ocr_data['width'][i], 'h': ocr_data['height'][i]})

        rows = {}
        for w in words:
            y_bin = round(w['top'] / 15) * 15
            if y_bin not in rows: rows[y_bin] = []
            rows[y_bin].append(w)

        in_box_13 = False 
        has_checked = False

        for y in sorted(rows.keys()):
            row_words = sorted(rows[y], key=lambda x: x['x0'])
            row_text = " ".join([w['text'] for w in row_words])

            if not in_box_13:
                if re.search(r'(?i)(13\.?|Where\s*appropriate)', row_text): in_box_13 = True
                continue 

            if in_box_13 and not has_checked and re.search(r'(?i)\b(Country|Invoicing|Back|CO)\b', row_text):
                # Neo vào chữ CO, Back, Country, hoặc Invoicing
                target_idx = -1
                for i, w in reversed(list(enumerate(row_words))):
                    if re.search(r'(?i)\b(CO|Back|Country|Invoicing)\b', w['text']):
                        target_idx = i
                        break
                
                if target_idx != -1:
                    anchor = row_words[target_idx]
                    h_word = anchor['h']
                    # Nội suy khoảng cách
                    box_size = int(h_word * 1.5)
                    x_box_end = anchor['x0'] - 10
                    if re.search(r'(?i)\bCO\b', anchor['text']): x_box_end -= int(13 * h_word * 0.45)
                    elif re.search(r'(?i)\bCountry\b', anchor['text']): x_box_end -= int(6 * h_word * 0.45)
                    elif re.search(r'(?i)\bInvoicing\b', anchor['text']): x_box_end -= int(14 * h_word * 0.45)
                    
                    x_box_start = max(0, x_box_end - box_size)
                    y_box_start = max(0, anchor['top'] - int((box_size - h_word) / 2))
                    y_box_end = y_box_start + box_size
                    
                    cb_img = img.crop((x_box_start, y_box_start, x_box_end, y_box_end))
                    gray_box = cb_img.convert("L")
                    pixels = list(gray_box.get_flattened_data()) if hasattr(gray_box, 'get_flattened_data') else list(gray_box.getdata())
                    dark_pixels = sum(1 for p in pixels if p < 128)
                    ratio = dark_pixels / len(pixels) if len(pixels) > 0 else 0
                    
                    if ratio > 0.035:
                        return "YES"
                    has_checked = True # Chỉ check 1 lần ô đầu tiên tìm thấy
    except Exception as e:
        print(f"Lỗi đọc Box 13 Scan: {e}")
    return "No"

def process_scanned_pdf(pdf, file_name):
    extracted_data = []
    page_first = pdf.pages[0]
    page_last = pdf.pages[-1]
    
    # 1. XỬ LÝ BOX 1 VÀ BOX 3 BẰNG TESSERACT THEO TỌA ĐỘ
    box1_img = page_first.crop((0, 0, 290, 110)).to_image(resolution=300).original
    box1_text = pytesseract.image_to_string(box1_img)
    exporter = re.sub(r'(?i).*?address, country\)', '', box1_text, flags=re.DOTALL).strip()
    
    box3_img = page_first.crop((0, 160, 300, 315)).to_image(resolution=300).original
    box3_text = pytesseract.image_to_string(box3_img)
    b3_clean = re.sub(r'(?i)Departure\s+Date\s*[:;]?', '', box3_text)
    b3_clean = re.sub(r'(?i)Vessel[\'’]?s?\s+Name.*?(?:\n|$)', '', b3_clean)
    b3_clean = re.sub(r'(?i)Port\s+of\s+Discharge\s*[:;]?', '', b3_clean)
    transport = re.sub(r'\s+', ' ', b3_clean).strip()
    
    # 2. XỬ LÝ REFERENCE NO VÀ FORM
    reference_no = file_name.replace(".pdf", "")
    
    docling_full_text = ""
    converter = DocumentConverter()
    for i, p in enumerate(pdf.pages):
        temp_path = f"temp_scanned_page_{i}.png"
        p.to_image(resolution=300).original.save(temp_path)
        res = converter.convert(temp_path)
        docling_full_text += "\n" + res.document.export_to_text()
        if os.path.exists(temp_path): os.remove(temp_path)
        
    form_type = ""
    form_match = re.search(r'(?i)FORM\s*([A-Za-z0-9]+)', docling_full_text)
    if form_match:
        form_type = form_match.group(1).upper()
        if form_type in ["AL", "A1", "A|", "A L"]: form_type = "AI"
        
    # 3. BOX 11 & 12
    produced_in = ""
    prod_match = re.search(r'(?i)produced in\s*\n*([A-Za-z\s]+)\s*\n*\s*\(Country\)', docling_full_text)
    if prod_match: produced_in = prod_match.group(1).strip().upper()
    
    exported_to = ""
    exp_match = re.search(r'(?i)exported to\s*\n*([A-Za-z\s]+)\s*\n*\s*\(Importing Country\)', docling_full_text)
    if exp_match:
        exported_to = re.sub(r'(?i)(for|Secret.*|Ministry.*)', '', exp_match.group(1)).strip().upper()
    
    date_cert = ""
    cert_match = re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s*\[', docling_full_text)
    if cert_match: date_cert = cert_match.group(1).strip()
    
    box_13_str = extract_box13_scanned(page_last)
    
    # 4. CHIA TÁCH VÀ LỌC CÁC ITEMS TỪ DOCLING RAW TEXT
    matches = list(re.finditer(r'\b(\d+)[\.\,]?\s+N/M', docling_full_text))
    
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(docling_full_text)
        block = docling_full_text[start:end]
        
        item_no = matches[i].group(1)
        
        carton = ""
        c_match = re.search(r'N/M\s+(\d+)\s*CARTON', block, re.IGNORECASE)
        if c_match: carton = c_match.group(1)
        
        origin, qty, uom = "", "", ""
        o_match = re.search(r'CARTON\s+(.*?)\s+(\d+)\s+([A-Za-z]+)', block, re.IGNORECASE)
        if o_match:
            origin = o_match.group(1).strip().upper()
            qty = o_match.group(2)
            uom = o_match.group(3).upper()
            
        usd = ""
        u_match = re.search(r'USD\s+([\d\.]+)', block)
        if u_match: usd = u_match.group(1)
        
        date_inv = ""
        d_match = re.search(r'(\d{2}/\d{2}/\d{4})', block)
        if d_match: date_inv = d_match.group(1)
        
        invoice = ""
        inv_match = re.search(r'(VN[A-Za-z0-9\-]+(?:\s+[A-Z]{2})?)', block)
        if inv_match:
            invoice = inv_match.group(1).strip()
            if form_type == "AI" and "IN" not in invoice:
                invoice += " IN"
                
        desc = ""
        if invoice:
            inv_esc = re.escape(invoice.split()[0])
            desc_match = re.search(f'{inv_esc}(?:\\s+[A-Z]{{2}})?\\s+(.*?)\\s+-', block)
            if desc_match: desc = desc_match.group(1).strip()
            
        imp_hs = ""
        imp_match = re.search(r'IMPORTING\s+COUNTRY\s+HS\s+CODE\s+(\d+)', block, re.IGNORECASE)
        if imp_match: imp_hs = imp_match.group(1)[:8]
        
        exp_hs = ""
        exp_match = re.search(r'EXPORTING\s+COUNTRY\s+HS\s+CODE\s+(\d+)', block, re.IGNORECASE)
        if exp_match: exp_hs = exp_match.group(1)[:8]
        
        orig_co = ""
        oc_match = re.search(r'Number:\s*([A-Za-z0-9/]+)', block, re.IGNORECASE)
        if oc_match: orig_co = oc_match.group(1)
        
        orig_date = ""
        od_match = re.search(r'Date:\s*(\d{1,2}-[A-Za-z]{3}-\d{4})', block, re.IGNORECASE)
        if od_match: orig_date = od_match.group(1)
        
        extracted_data.append({
            COLUMNS[0]: form_type,
            COLUMNS[1]: reference_no,
            COLUMNS[2]: clean_text(orig_co),
            COLUMNS[3]: item_no,
            COLUMNS[4]: clean_text(desc),
            COLUMNS[5]: clean_text(qty),
            COLUMNS[6]: clean_text(uom),
            COLUMNS[7]: usd,
            COLUMNS[8]: clean_text(origin),  
            COLUMNS[9]: imp_hs,
            COLUMNS[10]: exp_hs,      
            COLUMNS[11]: invoice,      
            COLUMNS[12]: date_inv,        
            COLUMNS[13]: clean_text(carton),
            COLUMNS[14]: clean_text(orig_date),
            COLUMNS[15]: "", 
            COLUMNS[16]: clean_text(date_cert),
            COLUMNS[17]: exporter,
            COLUMNS[18]: "", 
            COLUMNS[19]: transport,
            COLUMNS[20]: produced_in,
            COLUMNS[21]: exported_to,
            COLUMNS[22]: "N/M",
            COLUMNS[23]: box_13_str
        })
        
    return {"error": None, "data": extracted_data, "file_name": file_name}

# ==========================================
# 3. XỬ LÝ CHÍNH CHO FILE PDF CÓ TEXT LAYER (CODE CŨ GIỮ NGUYÊN)
# ==========================================
def extract_global_info(page_first, page_last):
    bbox_exporter = (30, 40, 290, 130)  
    bbox_consignee = (30, 130, 290, 200)
    bbox_transport = (30, 220, 290, 330)
    bbox_ref_no = (290, 40, 500, 110)

    exporter = clean_text(page_first.crop(bbox_exporter).extract_text())
    consignee = clean_text(page_first.crop(bbox_consignee).extract_text())
    transport = clean_text(page_first.crop(bbox_transport).extract_text())
    
    raw_ref = page_first.crop(bbox_ref_no).extract_text()
    ref_match = re.search(r'Reference No\.\s*([A-Z0-9\-]+)', raw_ref, re.IGNORECASE)
    reference_no = ref_match.group(1) if ref_match else clean_text(raw_ref)

    form_type = ""
    try:
        bbox_form = (290, 0, page_first.width, 150)
        cropped_form_page = page_first.crop(bbox_form)
        pil_image = cropped_form_page.to_image(resolution=200).original
        ocr_text = pytesseract.image_to_string(pil_image)
        form_match = re.search(r'FORM\s*([A-Za-z0-9]+)', ocr_text, re.IGNORECASE)
        if form_match: form_type = form_match.group(1).strip().upper()
    except:
        pass

    movement_cert = ""
    third_party = "No"
    try:
        bbox_box13 = (0, 600, page_last.width, page_last.height)
        cropped_box13 = page_last.crop(bbox_box13)
        img_box13 = cropped_box13.to_image(resolution=200).original
        ocr_data = pytesseract.image_to_data(img_box13, output_type=Output.DICT)
        
        def check_status(keyword_pattern):
            found_idx = -1
            for i, text in enumerate(ocr_data['text']):
                if re.search(keyword_pattern, text, re.IGNORECASE):
                    found_idx = i
                    break
            if found_idx == -1: return ""
            x = ocr_data['left'][found_idx] + 50 if 'Movement' in keyword_pattern or 'Back' in keyword_pattern else ocr_data['left'][found_idx]
            y = ocr_data['top'][found_idx]
            h = 30
            box_size = int(h * 1.5)
            x_box_start = max(0, x - box_size - int(h * 0.3))
            gray_box = img_box13.crop((x_box_start, max(0, y - int((box_size - h) / 2)), x - int(h * 0.3), max(0, y - int((box_size - h) / 2)) + box_size)).convert("L")
            pixels = list(gray_box.get_flattened_data()) if hasattr(gray_box, 'get_flattened_data') else list(gray_box.getdata())
            return "Yes" if pixels and sum(1 for p in pixels if p < 128) / len(pixels) > 0.12 else "No"

        if check_status('Movement') == "Yes": movement_cert = "Movement Certificate"
        elif check_status(r'Back-to-Back|Back') == "Yes": movement_cert = "Back-to-Back CO"
        if check_status('Third') == "Yes": third_party = "Yes"
    except:
        pass

    bbox_box11 = (0, 545, 350, page_last.height)
    bbox_box12 = (300, 550, page_last.width, page_last.height)
    box11_text = clean_text(page_last.crop(bbox_box11).extract_text())
    box12_text = clean_text(page_last.crop(bbox_box12).extract_text())
    
    asean_china = ["CHINA", "VIETNAM", "MALAYSIA", "SINGAPORE", "INDONESIA", "THAILAND", "PHILIPPINES", "BRUNEI", "CAMBODIA", "LAOS", "MYANMAR"]
    country_matches = re.findall(r'\b(' + '|'.join(asean_china) + r')\b', box11_text, re.IGNORECASE)
    produced_in = country_matches[0].upper() if len(country_matches) > 0 else ""
    exported_to = country_matches[1].upper() if len(country_matches) > 1 else ""
    
    date_cert = format_to_dd_mm_yyyy(re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', box12_text).group(1)) if re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', box12_text) else ""

    return {
        "exporter": exporter, "consignee": consignee, "transport": transport,
        "reference_no": reference_no, "movement_cert": movement_cert, "third_party": third_party,
        "produced_in": produced_in, "exported_to": exported_to, "date_of_cert": date_cert,
        "form_type": form_type
    }

def parse_description_fields(desc_text, weight_value_text=""):
    text = re.sub(r'\s+', ' ', desc_text).strip()
    weight_text = re.sub(r'\s+', ' ', weight_value_text).strip()
    
    carton_match = re.search(r'^([\d,\.]+)\s*CARTON', text, re.IGNORECASE)
    carton = carton_match.group(1).strip() if carton_match else ""
    
    qty, uom = "", ""
    if weight_text:
        weight_no_currency = re.sub(r'(?i)(USD|MYR|EUR|SGD|VND)\s*[\d,\.]+', '', weight_text).strip()
        qty_match = re.search(r'^([0-9][0-9,\.]*)', weight_no_currency)
        if qty_match: qty = qty_match.group(1).strip()
        uom_match = re.search(r'([A-Za-z]+)$', weight_no_currency)
        if uom_match: uom = uom_match.group(1).strip().upper()
            
    if not qty:
        num_of_match = re.search(r'(?:NUMBER|QUANTITY|AMOUNT|TOTAL)\s+OF\s+(PAIRS?|PIECES?|SETS?|PCE|PR|CARTONS?)\s*[-:]?\s*([\d,\.]+)', text, re.IGNORECASE)
        if num_of_match:
            uom = num_of_match.group(1).strip().upper()
            qty = num_of_match.group(2).strip()

    if not qty:
        qty_uom_match_7 = re.search(r'(?:-)?\s*([\d,\.]+)\s*([A-Za-z]+)\b', text, re.IGNORECASE)
        if qty_uom_match_7:
            qty = qty_uom_match_7.group(1).strip()
            uom = qty_uom_match_7.group(2).strip().upper()
            
    if uom and uom.endswith("S") and len(uom) > 3: uom = uom[:-1]

    desc_before_meta = re.split(r'(?i)(IMPORTING COUNTRY|EXPORTING COUNTRY|Original CO)', text)[0].strip()
    desc_cleaned = re.sub(r'^\s*[\d,\.]+\s*CARTONS?\s*(?:[-–—:]\s*)?', '', desc_before_meta, flags=re.IGNORECASE).strip()
    eng_desc = re.sub(r'(?:\s+[-–—:]\s+|\s+)([\d,\.]+)\s*[A-Za-z\s\.]*$', '', desc_cleaned).strip()
    eng_desc = re.sub(r'[-–—:,]\s*$', '', eng_desc).strip()
            
    import_hs = re.sub(r'\D', '', re.search(r'IMPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE).group(1))[:8] if re.search(r'IMPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE) else ""
    export_hs = re.sub(r'\D', '', re.search(r'EXPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE).group(1))[:8] if re.search(r'EXPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE) else ""
    orig_co = re.sub(r'[-–—:.,\s]+$', '', re.search(r'Original CO Reference Number\s*:\s*(.*?)(?=Issuance Date|Issuing Authority|TOTAL|$)', text, re.IGNORECASE).group(1).strip()) if re.search(r'Original CO Reference Number\s*:\s*(.*?)(?=Issuance Date|Issuing Authority|TOTAL|$)', text, re.IGNORECASE) else ""
    issue_date = re.search(r'Issuance Date\s*:\s*(.*?)(?=Issuing Authority|TOTAL|$)', text, re.IGNORECASE).group(1).strip() if re.search(r'Issuance Date\s*:\s*(.*?)(?=Issuing Authority|TOTAL|$)', text, re.IGNORECASE) else ""
    auth = re.search(r'Issuing Authority\s*:\s*(.*?)(?=TOTAL|Page|$)', text, re.IGNORECASE).group(1).strip() if re.search(r'Issuing Authority\s*:\s*(.*?)(?=TOTAL|Page|$)', text, re.IGNORECASE) else ""
    
    return carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth

def extract_table_items(pdf):
    items = []
    current_item = None
    global_invoice = ""
    third_party_text = ""

    for page_idx, page in enumerate(pdf.pages):
        table_crop = page.crop((0, 330, page.width, 555))
        words = table_crop.extract_words()
        
        rows = {}
        for w in words:
            y_bin = round(w['top'] / 5) * 5 
            if y_bin not in rows: rows[y_bin] = []
            rows[y_bin].append(w)
            
        page_changed = True 
        in_footer_section = False
        
        for y in sorted(rows.keys()):
            row_words = sorted(rows[y], key=lambda x: x['x0']) 
            row_text = " ".join([w['text'] for w in row_words])
            
            if "Item Number" in row_text or "Marks and" in row_text: continue
            if re.search(r'(?i)(Third\s+Party|\bTOTAL\b)', row_text): in_footer_section = True

            col_item_no = [w['text'] for w in row_words if 35 <= w['x0'] < 77]
            col_marks   = [w['text'] for w in row_words if 77 <= w['x0'] < 150]
            col_desc    = [w['text'] for w in row_words if 150 <= w['x0'] < 310]
            col_origin  = [w['text'] for w in row_words if 310 <= w['x0'] < 398]
            col_weight  = [w['text'] for w in row_words if 398 <= w['x0'] < 480]
            col_invoice = [w['text'] for w in row_words if w['x0'] >= 480]

            item_no_text = " ".join(col_item_no).strip()
            check_number = item_no_text.replace(".", "").replace(",", "").strip()

            if in_footer_section:
                if col_desc: third_party_text += " " + " ".join(col_desc)
                if col_invoice:
                    inv_text = " ".join(col_invoice)
                    if "VN" in inv_text and "/" in inv_text: global_invoice += " " + inv_text
                continue 

            if check_number.isdigit() and len(check_number) > 0: 
                if current_item: items.append(current_item) 
                current_item = {"item_no": item_no_text, "marks": " ".join(col_marks), "desc": " ".join(col_desc), "origin": " ".join(col_origin), "weight_value": " ".join(col_weight), "invoice": " ".join(col_invoice)}
                page_changed = False 
            elif current_item:
                if page_changed:
                    items.append(current_item) 
                    current_item = {"item_no": "CONTINUATION", "marks": "", "desc": "", "origin": "", "weight_value": "", "invoice": ""}
                    page_changed = False 
                if col_marks: current_item["marks"] += " " + " ".join(col_marks)
                if col_desc: current_item["desc"] += "\n" + " ".join(col_desc)
                if col_origin: current_item["origin"] += " " + " ".join(col_origin)
                if col_weight: current_item["weight_value"] += " " + " ".join(col_weight)
                if col_invoice: 
                    current_item["invoice"] += " " + " ".join(col_invoice)
                    if "VN" in current_item["invoice"] and "/" in current_item["invoice"]: global_invoice = current_item["invoice"]

    if current_item: items.append(current_item)
    third_party_text = re.split(r'(?i)TOTAL|USD|MYR|EUR', clean_text(third_party_text))[0].strip()
    return items, global_invoice, third_party_text

def process_single_pdf(file_data):
    file_name = file_data["name"]
    file_bytes = file_data["bytes"]
    extracted_data = []
    
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            # KIỂM TRA TEXT LAYER - Nếu < 50 ký tự -> File Scan
            first_page_text = pdf.pages[0].extract_text()
            if not first_page_text or len(first_page_text.strip()) < 50:
                if "scanned_files_detected" in st.session_state:
                    st.session_state.scanned_files_detected.append(file_name)
                return process_scanned_pdf(pdf, file_name)

            # LUỒNG CHO FILE CÓ TEXT LAYER
            global_info = extract_global_info(pdf.pages[0], pdf.pages[-1])
            items, global_invoice, third_party_column_val = extract_table_items(pdf)
            box_13_str = "YES" if (global_info["third_party"] == "Yes" and global_info["movement_cert"] != "") else "No"

            final_items = []
            for item in items:
                if item["item_no"] == "CONTINUATION" or not item["item_no"].strip():
                    if final_items: 
                        final_items[-1]["marks"] = (final_items[-1]["marks"] + " " + item["marks"]).strip()
                        final_items[-1]["desc"] = (final_items[-1]["desc"] + "\n" + item["desc"]).strip()
                        final_items[-1]["origin"] = (final_items[-1]["origin"] + " " + item["origin"]).strip()
                        final_items[-1]["weight_value"] = (final_items[-1]["weight_value"] + " " + item["weight_value"]).strip()
                        final_items[-1]["invoice"] = (final_items[-1]["invoice"] + " " + item["invoice"]).strip()
                else:
                    final_items.append(item)
            
            for item in final_items:
                desc = clean_text(item["desc"])
                invoice = clean_text(item["invoice"]) if item["invoice"] else global_invoice
                weight_value_text = clean_text(item["weight_value"])
                
                invoice_number, invoice_date = "", ""
                if invoice:
                    date_match = re.search(r'(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})', invoice)
                    if date_match:
                        invoice_date = format_to_dd_mm_yyyy(date_match.group(1))
                        invoice_number = invoice.replace(date_match.group(1), "").strip()
                    else: invoice_number = invoice
                
                carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth = parse_description_fields(desc, weight_value_text)
                
                usd_match = re.search(r'USD\s*([\d,\.]+)', weight_value_text, re.IGNORECASE)
                usd = usd_match.group(1).strip() if usd_match else ""
                if not usd:
                    usd_match_desc = re.search(r'USD\s*([\d,\.]+)', desc, re.IGNORECASE)
                    usd = usd_match_desc.group(1).strip() if usd_match_desc else ""
                
                item_no_val = clean_text(item["item_no"])
                if item_no_val.upper() == "CONTINUATION": item_no_val = ""
                
                extracted_data.append({
                    COLUMNS[0]: global_info["form_type"], COLUMNS[1]: global_info["reference_no"],
                    COLUMNS[2]: clean_text(orig_co), COLUMNS[3]: item_no_val, COLUMNS[4]: clean_text(eng_desc),
                    COLUMNS[5]: clean_text(qty), COLUMNS[6]: clean_text(uom), COLUMNS[7]: usd,
                    COLUMNS[8]: clean_text(item["origin"]), COLUMNS[9]: import_hs, COLUMNS[10]: export_hs,      
                    COLUMNS[11]: invoice_number, COLUMNS[12]: invoice_date, COLUMNS[13]: clean_text(carton),
                    COLUMNS[14]: format_to_dd_mm_yyyy(issue_date), COLUMNS[15]: clean_text(auth),
                    COLUMNS[16]: global_info["date_of_cert"], COLUMNS[17]: global_info["exporter"],
                    COLUMNS[18]: global_info["consignee"], COLUMNS[19]: global_info["transport"],
                    COLUMNS[20]: global_info["produced_in"], COLUMNS[21]: global_info["exported_to"],
                    COLUMNS[22]: clean_text(item["marks"]), COLUMNS[23]: box_13_str
                })
    except Exception as e:
        return {"error": f"{file_name}: Lỗi trích xuất - {str(e)}", "data": [], "file_name": file_name}
    return {"error": None, "data": extracted_data, "file_name": file_name}

# ==========================================
# 4. GIAO DIỆN STREAMLIT
# ==========================================
def init_session_state():
    if "pdf_files" not in st.session_state: st.session_state.pdf_files = []
    if "is_processing" not in st.session_state: st.session_state.is_processing = False
    if "extracted_data" not in st.session_state: st.session_state.extracted_data = None
    if "errors" not in st.session_state: st.session_state.errors = []
    if "scanned_files_detected" not in st.session_state: st.session_state.scanned_files_detected = []

def reset_data_state():
    st.session_state.extracted_data = None
    st.session_state.errors = []
    st.session_state.scanned_files_detected = []

def main():
    st.set_page_config(page_title="Form E Extractor", layout="wide", page_icon="📑")
    init_session_state()
    
    if "has_sent_email" not in st.session_state:
        send_email_notification()
        st.session_state.has_sent_email = True
    
    st.markdown(f"""
        <style>
        .stApp {{ background-color: {BG_LIGHT}; font-family: 'Inter', sans-serif; }}
        #MainMenu, footer, header {{ visibility: hidden; }}
        .app-header {{ padding: 1.5rem 0 2rem 0; text-align: center; }}
        .app-title {{ color: {DECATHLON_DARK}; font-weight: 800; font-size: 2.5rem; letter-spacing: -1px; margin-bottom: 0.5rem; }}
        .highlight {{ color: {DECATHLON_BLUE}; }}
        .control-panel {{ background: white; padding: 2rem; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); border: 1px solid #E5E7EB; margin-bottom: 2rem; }}
        button[kind="primary"] {{ background-color: {DECATHLON_BLUE}; color: white; border: none; height: 48px !important; border-radius: 8px !important; font-weight: 600 !important; width: 100% !important; }}
        button[kind="primary"]:disabled {{ background-color: #D1D5DB; color: #9CA3AF; }}
        button[kind="secondary"] {{ border: 1px solid #D1D5DB; height: 48px !important; border-radius: 8px !important; font-weight: 600 !important; width: 100% !important; }}
        .stProgress > div > div > div > div {{ background-color: {DECATHLON_BLUE}; border-radius: 10px; height: 8px; }}
        .data-card {{ background: white; padding: 1.5rem; border-radius: 16px; border: 1px solid #E5E7EB; }}
        .badge {{ display: inline-block; padding: 0.35rem 0.8rem; font-size: 0.85rem; font-weight: 600; border-radius: 9999px; background-color: #E0F2FE; color: #0369A1; margin-bottom: 1rem; }}
        </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
        <div class="app-header">
            <h1 style="font-size: 3.5rem; margin-bottom: 0;">📑➔📊</h1>
            <div class="app-title"><span class="highlight">PDF</span> Extract to <span class="highlight">Excel</span> Pipeline</div>
            <div style="color: #6B7280;">Hệ thống trích xuất PDF tự động (Hỗ trợ cả file Scan)</div>
        </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="control-panel">', unsafe_allow_html=True)
        st.markdown("<div class='badge'>⚙️ BẢNG ĐIỀU KHIỂN</div>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3, gap="large")
        cancel_clicked = False

        with col1:
            st.markdown("**1. Nguồn dữ liệu PDF**")
            uploaded_files = st.file_uploader("Kéo thả File vào đây", type=["pdf"], accept_multiple_files=True, disabled=st.session_state.is_processing)
            if uploaded_files:
                st.session_state.pdf_files = uploaded_files
                with st.expander("🛠️ Quản lý tệp (Bấm để xem)", expanded=True):
                    with st.container(height=200):
                        for f in st.session_state.pdf_files: st.markdown(f"📄 `{f.name}`")
            else: st.session_state.pdf_files = []

        with col2:
            st.markdown("**2. Quá trình xử lý**")
            has_files = len(st.session_state.pdf_files) > 0
            if has_files:
                if not st.session_state.is_processing:
                    if st.button("🚀 BẮT ĐẦU TRÍCH XUẤT", type="primary"):
                        st.session_state.is_processing = True
                        reset_data_state()
                        st.rerun()
                else:
                    if st.button("🛑 Hủy tiến trình", type="secondary"): cancel_clicked = True
            else:
                st.button("🚀 BẮT ĐẦU TRÍCH XUẤT", type="primary", disabled=True)

        with col3:
            st.markdown("**3. Xuất kết quả**")
            if st.session_state.extracted_data is not None:
                df_result = pd.DataFrame(st.session_state.extracted_data, columns=COLUMNS)
                numeric_cols = ["Item Number", "Quantity", "USD", "IMPORTING COUNTRY HS CODE", "EXPORTING COUNTRY HS CODE", "CARTON"]
                for col in numeric_cols:
                    if col in df_result.columns:
                        df_result[col] = df_result[col].astype(str).str.replace(',', '')
                        df_result[col] = pd.to_numeric(df_result[col], errors='coerce')
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="C_O_FormE")
                
                st.download_button(label="📥 TẢI FILE EXCEL (.XLSX)", data=output.getvalue(), file_name="Decathlon_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            else: st.button("📥 TẢI FILE EXCEL", disabled=True)

        if cancel_clicked:
            countdown = st.empty()
            for i in range(3, 0, -1):
                countdown.error(f"⚠️ Đã ngắt tiến trình. Đang dọn dẹp... {i}s")
                time.sleep(1)
            reset_data_state()
            st.session_state.is_processing = False
            st.session_state.pdf_files = []
            st.rerun()
            
        elif st.session_state.is_processing:
            st.markdown("<hr>", unsafe_allow_html=True)
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            all_extracted_data = []
            errors = []
            total_files = len(st.session_state.pdf_files)
            
            safe_file_list = [{"name": f.name, "bytes": f.getvalue()} for f in st.session_state.pdf_files]
            
            processed_count = 0
            for f_data in safe_file_list:
                processed_count += 1
                result = process_single_pdf(f_data)
                
                if result["error"]: errors.append(result["error"])
                else: all_extracted_data.extend(result["data"])
                
                progress_bar.progress(processed_count / total_files)
                status_text.info(f"⏳ Đã xử lý: `{result['file_name']}` ({processed_count}/{total_files})")
                
                del result
                gc.collect()

            st.session_state.extracted_data = all_extracted_data
            st.session_state.errors = errors
            st.session_state.is_processing = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True) 

    if st.session_state.extracted_data is not None:
        # CẢNH BÁO UI CHO FILE SCAN
        if st.session_state.scanned_files_detected:
            st.info("ℹ️ **Hệ thống đã nhận diện và áp dụng thuật toán ĐẶC BIỆT cho các file SCAN sau:**\n\n" + 
                    "\n".join([f"- {f}" for f in st.session_state.scanned_files_detected]))

        if st.session_state.errors:
            with st.expander("⚠️ Có lỗi xảy ra (Click để xem chi tiết)"):
                for err in st.session_state.errors: st.error(err)

        if st.session_state.extracted_data:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            st.markdown(f"#### 📊 Dữ liệu chi tiết ({len(st.session_state.extracted_data)} dòng)")
            df_result = pd.DataFrame(st.session_state.extracted_data, columns=COLUMNS)
            st.dataframe(df_result, width='stretch', height=450)
            st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
