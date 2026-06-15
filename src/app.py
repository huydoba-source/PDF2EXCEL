import gc
import streamlit as st
import pandas as pd
import os
import io
import re
import time
import requests  
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- THƯ VIỆN BỔ SUNG CHO OCR ---
import pdfplumber
import fitz  # PyMuPDF
import pytesseract
from pytesseract import Output
from PIL import Image, ImageDraw, ImageOps

# [LƯU Ý]: Nếu chạy trên máy tính Windows thì đổi thành r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

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
# XỬ LÝ BOX 13 SCAN
# ==========================================
def extract_box13_scanned(page):
    try:
        bbox_13 = (0, page.height - 250, page.width, page.height)
        crop = page.crop(bbox_13)
        img = crop.to_image(resolution=300).original
        
        ocr_data = pytesseract.image_to_data(img, output_type=Output.DICT, config='--oem 3 --psm 11')
        words = []
        for i in range(len(ocr_data['text'])):
            text = ocr_data['text'][i].strip()
            if not text or len(text) < 2: continue
            words.append({
                'text': text, 'x0': ocr_data['left'][i], 'top': ocr_data['top'][i],
                'x1': ocr_data['left'][i] + ocr_data['width'][i],
                'bottom': ocr_data['top'][i] + ocr_data['height'][i], 'h': ocr_data['height'][i]
            })

        rows = {}
        for w in words:
            y_bin = round(w['top'] / 15) * 15
            if y_bin not in rows: rows[y_bin] = []
            rows[y_bin].append(w)

        results = {"Third Country Invoicing": "No", "Back-to-Back CO": "No"}
        in_box_13 = False
        third_done = False
        back_done = False

        for y in sorted(rows.keys()):
            row_words = sorted(rows[y], key=lambda x: x['x0'])
            row_text = " ".join([w['text'] for w in row_words])

            if not in_box_13:
                if re.search(r'(?i)(13\.?|Where\s*appropriate)', row_text): in_box_13 = True
                else: continue

            if in_box_13 and not third_done and re.search(r'(?i)\b(Country|Invoicing)\b', row_text):
                target_idx = -1
                for i, w in enumerate(row_words):
                    if re.search(r'(?i)\b(Country|Invoicing)\b', w['text']):
                        target_idx = i; break
                if target_idx != -1:
                    start_word = row_words[target_idx]
                    for i in range(target_idx - 1, -1, -1):
                        gap = row_words[i+1]['x0'] - row_words[i]['x1']
                        if gap < 80 and re.search(r'(?i)(Third|Count|Invoic|hir|oun)', row_words[i]['text']):
                            start_word = row_words[i]
                        else: break
                    x_box_end = start_word['x0'] - 8
                    if not re.search(r'(?i)(Third|hir)', start_word['text']): x_box_end -= 60
                    h_word = start_word['h']
                    box_size = int(60 * 1.4) 
                    x_box_start = max(0, x_box_end - box_size)
                    y_box_start = max(0, start_word['top'] - int((box_size - h_word) / 2))
                    y_box_end = y_box_start + box_size

                    cb_img = img.crop((x_box_start, y_box_start, x_box_end, y_box_end))
                    gray_box = cb_img.convert("L")
                    pixels = list(gray_box.get_flattened_data()) if hasattr(gray_box, 'get_flattened_data') else list(gray_box.getdata())
                    dark_pixels = sum(1 for p in pixels if p < 128)
                    if (dark_pixels / len(pixels) if len(pixels) > 0 else 0) > 0.035:
                        results["Third Country Invoicing"] = "YES"
                    third_done = True

            if in_box_13 and not back_done and re.search(r'(?i)\b(Back|CO)\b', row_text):
                target_idx = -1
                for i, w in reversed(list(enumerate(row_words))):
                    if re.search(r'(?i)\b(CO|Back)\b', w['text']):
                        target_idx = i; break
                if target_idx != -1:
                    start_word = row_words[target_idx]
                    for i in range(target_idx - 1, -1, -1):
                        gap = row_words[i+1]['x0'] - row_words[i]['x1']
                        if gap < 80 and re.search(r'(?i)(Back|ack|to|\-)', row_words[i]['text']):
                            start_word = row_words[i]
                        else: break
                    start_text = start_word['text']
                    box_size = int(60 * 1.4) 
                    if re.search(r'(?i)^CO$', start_text):
                        x_box_end = start_word['x0'] - 140
                        x_box_start = max(0, x_box_end - box_size)
                    else:
                        match = re.search(r'(?i)back', start_text)
                        if match and match.start() >= 2:
                            x_box_start = start_word['x0']
                            x_box_end = x_box_start + box_size
                        else:
                            x_box_end = start_word['x0'] - 10
                            x_box_start = max(0, x_box_end - box_size)
                    h_word = start_word['h']
                    y_box_start = max(0, start_word['top'] - int((box_size - h_word) / 2))
                    y_box_end = y_box_start + box_size

                    cb_img = img.crop((x_box_start, y_box_start, x_box_end, y_box_end))
                    gray_box = cb_img.convert("L")
                    pixels = list(gray_box.get_flattened_data()) if hasattr(gray_box, 'get_flattened_data') else list(gray_box.getdata())
                    dark_pixels = sum(1 for p in pixels if p < 128)
                    if (dark_pixels / len(pixels) if len(pixels) > 0 else 0) > 0.035:
                        results["Back-to-Back CO"] = "YES"
                    back_done = True 

        if results["Third Country Invoicing"] == "YES" or results["Back-to-Back CO"] == "YES": return "YES"
        return "No"
    except Exception as e:
        print(f"Lỗi đọc Box 13 Scan: {e}")
        return "No"

# ==========================================
# 2. XỬ LÝ FILE SCAN BẰNG FITZ (PyMuPDF) + TESSERACT THEO TỌA ĐỘ
# ==========================================
def process_scanned_pdf(pdf, file_bytes, file_name):
    extracted_data = []
    
    # 1. HARDCODE BOX 1 & BOX 2
    exporter = "DECATHLON LOGISTICS MALAYSIA SDN. BHD.\nPLOT D40 & D44\nJALAN DPB/8, ZONE B\nPELABUHAN TANJUNG PELEPAS\n81560 GELANG PATAH, JOHOR, MALAYSIA"
    consignee = "DECATHLON VIETNAM CO., LTD\nPAX SKY BUILDING, 5TH FLOOR\n26 UNG VAN KHIEM, WARD 25,\nBINH THANH DISTRICT\n700000 HO CHI MINH CITY VIETNAM"
    
    reference_no = file_name.replace(".pdf", "")
    
    # DÙNG FITZ ĐỂ MỞ PDF 
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        page_first = doc[0]
        
        # ZOOM chuẩn 300 DPI cho OCR Tesseract
        zoom = 300 / 72
        mat = fitz.Matrix(zoom, zoom)
        tess_config = r'--oem 3 --psm 6'

        # ---------------------------------------------------------
        # LẤY THÔNG TIN GLOBAL TỪ TRANG ĐẦU (Tọa độ: Top)
        # ---------------------------------------------------------
        top_rect = fitz.Rect(0, 0, page_first.rect.width, 300)
        pix_top = page_first.get_pixmap(matrix=mat, clip=top_rect)
        img_top = Image.frombytes("RGB", [pix_top.width, pix_top.height], pix_top.samples) if not pix_top.alpha else Image.frombytes("RGBA", [pix_top.width, pix_top.height], pix_top.samples).convert("RGB")
        top_text = pytesseract.image_to_string(img_top, config=tess_config).strip()
        top_flat = re.sub(r'[\n\|]', ' ', top_text)

        form_type = ""
        form_match = re.search(r'(?i)FORM\s*([A-Za-z0-9]+)', top_flat)
        if form_match:
            form_type = form_match.group(1).upper()
            if form_type in ["AL", "A1", "A|", "A L", "A I"]: form_type = "AI"

        produced_in, exported_to = "", ""
        if form_type == "AI":
            produced_in = "INDIA"
            exported_to = "VIETNAM"
        else:
            prod_match = re.search(r'([A-Za-z\s]+)\s*\(Country\)', top_flat, re.IGNORECASE)
            if prod_match: produced_in = prod_match.group(1).strip().upper()
            exp_to_match = re.search(r'exported to(.*?)\(Importing Country\)', top_flat, re.IGNORECASE)
            if exp_to_match:
                raw_exp = exp_to_match.group(1)
                caps = re.findall(r'\b[A-Z]{3,}\b', raw_exp)
                exported_to = caps[-1] if caps else raw_exp.strip().upper()

        # ---------------------------------------------------------
        # TRÍCH XUẤT BOX 3 (XÓA TIÊU ĐỀ, TÌM NGÀY THÁNG)
        # ---------------------------------------------------------
        box3_rect = fitz.Rect(0, 160, 300, 315)
        pix3 = page_first.get_pixmap(matrix=mat, clip=box3_rect)
        img3 = Image.frombytes("RGB", [pix3.width, pix3.height], pix3.samples) if not pix3.alpha else Image.frombytes("RGBA", [pix3.width, pix3.height], pix3.samples).convert("RGB")
        
        b3_clean = pytesseract.image_to_string(img3, config=tess_config).strip()
        b3_clean = re.sub(r'\s+', ' ', b3_clean).strip()
        b3_clean = re.sub(r'(?i)3\.?\s*Means of transport.*?\)', '', b3_clean)
        b3_clean = re.sub(r'(?i)Departure Date\s*[:;]?', '', b3_clean)
        b3_clean = re.sub(r'(?i)Vessel[\'’]?s?\s*Name[/\\]?Aircraft.*?etc\.?[:;]?', '', b3_clean)
        b3_clean = re.sub(r'(?i)Port\s+of\s+Discharge\s*[:;]?', '', b3_clean)
        
        transport = ""
        transport_match = re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4}.*)', b3_clean, re.IGNORECASE)
        if transport_match: transport = transport_match.group(1).strip()
        else: transport = b3_clean.strip()
        transport = re.sub(r'\s+', ' ', transport).strip()

        # ---------------------------------------------------------
        # TRÍCH XUẤT DATE OF CERTIFICATION
        # ---------------------------------------------------------
        cert_rect = fitz.Rect(0, 550, page_first.rect.width, page_first.rect.height)
        pix_cert = page_first.get_pixmap(matrix=mat, clip=cert_rect)
        img_cert = Image.frombytes("RGB", [pix_cert.width, pix_cert.height], pix_cert.samples) if not pix_cert.alpha else Image.frombytes("RGBA", [pix_cert.width, pix_cert.height], pix_cert.samples).convert("RGB")
        cert_text = pytesseract.image_to_string(img_cert, config=tess_config).strip()
        
        date_cert = ""
        cert_match = re.search(r'Lumpur\s*,\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})', cert_text, re.IGNORECASE)
        if cert_match:
            date_cert = cert_match.group(1).strip()

        # ---------------------------------------------------------
        # TRÍCH XUẤT VÙNG HÀNG HÓA TỪ TẤT CẢ CÁC TRANG 
        # ---------------------------------------------------------
        main_text = ""
        for page in doc:
            x1 = 0
            y1 = 300
            x2 = page.rect.width
            y2 = page.rect.height - 180
            if y2 <= y1: continue
            
            main_rect = fitz.Rect(x1, y1, x2, y2)
            pix_main = page.get_pixmap(matrix=mat, clip=main_rect)
            img_main = Image.frombytes("RGB", [pix_main.width, pix_main.height], pix_main.samples) if not pix_main.alpha else Image.frombytes("RGBA", [pix_main.width, pix_main.height], pix_main.samples).convert("RGB")
            main_text += "\n" + pytesseract.image_to_string(img_main, config=tess_config)
            
    # Box 13 xử lý bằng hàm phụ (cũ)
    box_13_str = extract_box13_scanned(pdf.pages[-1])

    # ---------------------------------------------------------
    # 5. CẮT BLOCK VÀ ÁP DỤNG LOGIC CHỐNG NHIỄU OCR (MỚI)
    # ---------------------------------------------------------
    matches = list(re.finditer(r'(?i)(?:N/M|N\s*/\s*M|N/W|M/N|N\.M)', main_text))
    
    if not matches:
        extracted_data.append({
            COLUMNS[0]: form_type, COLUMNS[1]: reference_no, COLUMNS[2]: "", COLUMNS[3]: "",
            COLUMNS[4]: "[Lỗi OCR] Không tìm thấy dữ liệu Item.", COLUMNS[5]: "", COLUMNS[6]: "", 
            COLUMNS[7]: "", COLUMNS[8]: "", COLUMNS[9]: "", COLUMNS[10]: "", COLUMNS[11]: "", 
            COLUMNS[12]: "", COLUMNS[13]: "", COLUMNS[14]: "", COLUMNS[15]: "", COLUMNS[16]: date_cert, 
            COLUMNS[17]: exporter, COLUMNS[18]: consignee, COLUMNS[19]: transport, COLUMNS[20]: produced_in, 
            COLUMNS[21]: exported_to, COLUMNS[22]: "", COLUMNS[23]: box_13_str
        })
    else:
        for i in range(len(matches)):
            start = matches[i].start()
            end = matches[i+1].start() if i + 1 < len(matches) else len(main_text)
            block = main_text[start:end]
            
            item_no = str(i + 1)
            
            # --- LOGIC 1: BẮT MỎ NEO "USD" VÀ "DATE" ---
            qty, uom, usd, date_inv, origin_extra = "", "", "", "", ""
            usd_start_idx = -1
            
            usd_pattern = r'(\d[\d\.,]*)\s+([A-Za-z]{2,})(.*?)\s*USD\s+([\d\.,]+)\s+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})'
            usd_match = re.search(usd_pattern, block, re.IGNORECASE)
            
            if usd_match:
                qty = usd_match.group(1).strip()
                uom = usd_match.group(2).strip().upper()
                origin_extra = usd_match.group(3).strip()
                usd = usd_match.group(4).strip()
                date_inv = usd_match.group(5).strip()
                usd_start_idx = usd_match.start() # Lấy vị trí để tìm Description

            # --- LOGIC 2: BẮT MỎ NEO "CARTON" ĐỂ LẤY CARTON, ORIGIN VÀ INVOICE ---
            # Bỏ qua chữ cái bị thừa giữa Số và CARTON (Ví dụ: 1S CARTON -> Lấy 1)
            c_match = re.search(r'(\d+)[A-Za-z]*\s*CARTON', block, re.IGNORECASE)
            carton = c_match.group(1) if c_match else ""
            
            origin, invoice, invoice_raw = "", "", ""
            inv_end_idx = -1
            
            carton_line_match = re.search(r'(?m)^.*?(?:N/M|N\s*/\s*M).*?\d+[A-Za-z]*\s*CARTON\s+(.*?)$', block, re.IGNORECASE)
            if carton_line_match:
                rest_of_line = carton_line_match.group(1).strip()
                
                # Tìm mã Invoice (Bắt lỗi OCR JN, [N, |IN..., và bắt cả YN thay vì VN)
                vn_match = re.search(r'([VY]N[A-Z0-9\-]+(?:\s*(?:IN|JN|I\}|\|IN|\[N))?)', rest_of_line, re.IGNORECASE)
                if vn_match:
                    invoice_raw = vn_match.group(1)
                    base_inv = re.search(r'([VY]N[A-Z0-9\-]+)', invoice_raw, re.IGNORECASE).group(1).upper()
                    
                    # Sửa lỗi YN -> VN
                    if base_inv.startswith("YN"):
                        base_inv = "VN" + base_inv[2:]
                        
                    # Chuẩn hóa đuôi IN
                    invoice_clean = re.sub(r'(?i)(JN|I\}|\|IN|\[N)$', 'IN', invoice_raw).upper()
                    if invoice_clean.startswith("YN"):
                        invoice_clean = "VN" + invoice_clean[2:]
                        
                    if form_type == "AI" and not invoice_clean.endswith("IN"):
                        invoice = base_inv + " IN"
                    else:
                        invoice = invoice_clean
                        
                    # Lưu lại vị trí để cắt Description
                    inv_end_idx = block.find(vn_match.group(1)) + len(vn_match.group(1))
                    
                    # Origin là tất cả text nằm trước cụm (Số + Đơn vị tính)
                    before_inv = rest_of_line[:vn_match.start()].strip()
                    origin_split = re.search(r'^(.*?)(?=\s*\d+[\s\n]*[A-Za-z]+$)', before_inv)
                    if origin_split:
                        origin = origin_split.group(1).strip().upper()
                    else:
                        origin = before_inv.upper()
            
            # Gộp Origin bị rớt dòng (Ví dụ CTSH bị rớt xuống cùng dòng USD)
            if origin_extra: 
                origin = (origin + " " + re.sub(r'[-–—~_]+', '', origin_extra)).strip().upper()

            # --- LOGIC 3: LẤY ĐÚNG DESCRIPTION KẸP GIỮA INVOICE VÀ USD ---
            desc = ""
            if inv_end_idx != -1 and usd_start_idx != -1 and usd_start_idx > inv_end_idx:
                desc_raw = block[inv_end_idx:usd_start_idx]
                desc_raw = re.sub(r'^[\s\n]+', '', desc_raw) # Cắt khoảng trắng/xuống dòng thừa ở đầu
                desc_raw = re.sub(r'[-–—~_\s\n]+$', '', desc_raw) # Cắt bỏ rác và dấu trừ ở đuôi
                desc = clean_text(desc_raw)
            elif usd_start_idx != -1:
                # Fallback nếu không tìm thấy Invoice Number
                desc_fallback = block[:usd_start_idx].split('\n')[-1]
                desc = re.sub(r'[-–—~_]+$', '', desc_fallback).strip()

            # --- LOGIC 4: CẬP NHẬT TÌM HS CODES & ORIGINAL CO (CHỐNG NHIỄU OCR) ---
            imp_match = re.search(r'(?i)IMPORTING\s+COUNTRY\s+HS\s+CODE\s+(\d{10})', block)
            imp_hs = imp_match.group(1)[:8] if imp_match else ""
            
            exp_match = re.search(r'(?i)EXPORTING\s+COUNTRY\s+HS\s+CODE\s+(\d{10})', block)
            exp_hs = exp_match.group(1)[:8] if exp_match else ""
            
            # Khắc phục lỗi chữ "Onginal" hoặc "Orginal"
            orig_match = re.search(r'(?i)CO\s+Reference\s+Number:\s*\n*([A-Za-z0-9/]+)', block)
            orig_co = orig_match.group(1) if orig_match else ""
            
            iss_match = re.search(r'(?i)Issuance\s+Date:\s*\n*(\d{1,2}-[A-Za-z]{3}-\d{4})', block)
            orig_date = iss_match.group(1).upper() if iss_match else ""
            
            # LƯU KẾT QUẢ VÀO DỮ LIỆU
            extracted_data.append({
                COLUMNS[0]: form_type, COLUMNS[1]: reference_no, COLUMNS[2]: clean_text(orig_co),
                COLUMNS[3]: item_no, COLUMNS[4]: desc, COLUMNS[5]: clean_text(qty),
                COLUMNS[6]: clean_text(uom), COLUMNS[7]: usd, COLUMNS[8]: origin,  
                COLUMNS[9]: imp_hs, COLUMNS[10]: exp_hs, COLUMNS[11]: invoice,      
                COLUMNS[12]: date_inv, COLUMNS[13]: clean_text(carton), COLUMNS[14]: clean_text(orig_date),
                COLUMNS[15]: "", COLUMNS[16]: date_cert, COLUMNS[17]: exporter,
                COLUMNS[18]: consignee, COLUMNS[19]: transport, COLUMNS[20]: produced_in,
                COLUMNS[21]: exported_to, COLUMNS[22]: "N/M", COLUMNS[23]: box_13_str
            })
            
    # ==========================================
    # ĐOẠN CODE IN LOG KIỂM TRA
    # ==========================================
    print(f"\n" + "="*70)
    print(f"📊 KẾT QUẢ TRÍCH XUẤT (PyMuPDF + Tesseract) FILE: {file_name}")
    print("="*70)
    
    if not extracted_data:
        print("[!] Không có dữ liệu nào được trích xuất.")
    else:
        for idx, row in enumerate(extracted_data):
            item_id = row.get("Item Number", "N/A")
            print(f"\n📦 --- Dòng hàng hóa thứ {idx + 1} (Item No: {item_id}) ---")
            for key, value in row.items():
                if value and str(value).strip() != "":
                    display_value = str(value)
                    if len(display_value) > 100: display_value = display_value[:97] + "..."
                    print(f"   + {key}: {display_value}")
    print("="*70 + "\n")

    return {"error": None, "data": extracted_data, "file_name": file_name}

# ==========================================
# 3. XỬ LÝ CHÍNH CHO FILE PDF CÓ TEXT LAYER (CŨ)
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
    except: pass

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
    except: pass

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
            # KIỂM TRA TEXT LAYER 
            first_page_text = pdf.pages[0].extract_text()
            if not first_page_text or len(first_page_text.strip()) < 50:
                if "scanned_files_detected" in st.session_state:
                    st.session_state.scanned_files_detected.append(file_name)
                # Dùng Tesseract và tọa độ trực tiếp cho file scan
                return process_scanned_pdf(pdf, file_bytes, file_name)

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
