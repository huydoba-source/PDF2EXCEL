import gc
import streamlit as st
import pandas as pd
import os
import io
import re
import time
import requests  # Bổ sung thư viện requests thay cho smtplib
import pdfplumber
from concurrent.futures import ThreadPoolExecutor, as_completed
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# --- THƯ VIỆN BỔ SUNG CHO OCR ---
import pytesseract
from pytesseract import Output
from PIL import Image, ImageDraw

# [LƯU Ý]: Code tĩnh của tesseract_cmd đã được vô hiệu hóa để có thể chạy trên Web/Docker.
# Nếu bạn muốn test trực tiếp trên môi trường Windows local của bạn, hãy bỏ comment dòng dưới:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ==========================================
# 1. ĐỒNG BỘ CẤU HÌNH CỘT DỮ LIỆU ĐẦU RA MỚI KỲ VỌNG
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
    "Gross weight or net weight or other quantity, and value",
    "Original CO Issuance Date",
    "Issuing Authority",
    "Number and type of packages, description of products",
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
# HÀM GỬI EMAIL KHÔNG CẦN MẬT KHẨU (SỬ DỤNG FORMSUBMIT API)
# ==========================================

def send_email_notification():
    # 1. EMAIL TRẠM PHÁT (Hãy tạo 1 cái Gmail ảo và tạo Mật khẩu ứng dụng 16 ký tự)
    SENDER_EMAIL = "dobahuy7@gmail.com" 
    SENDER_PASSWORD = "kwyv yjud qvhy ehiq" 
    
    # 2. EMAIL NHẬN THÔNG BÁO CỦA BẠN
    RECEIVER_EMAIL = "huy.doba@decathlon.com" 
    
    # Chặn chạy nếu chưa cài đặt email trạm
    if "nhap_gmail_ao" in SENDER_EMAIL:
        return

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
        print("✅ Đã gửi email thông báo thành công!")
    except Exception as e:
        print(f"[!] Lỗi khi gửi email SMTP: {e}")

# ==========================================
# 2. HÀM CHUẨN HÓA ĐỊNH DẠNG NGÀY THÁNG
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
    if match:
        d, m, y = match.group(1), match.group(2), match.group(3)
        return f"{int(d):02d}-{int(m):02d}-{y}"
        
    match = re.search(r'(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})', date_str)
    if match:
        y, m, d = match.group(1), match.group(2), match.group(3)
        return f"{int(d):02d}-{int(m):02d}-{y}"

    match = re.search(r'(\d{1,2})\s*[\-\s]\s*([A-Za-z]+)\s*[\-\s]\s*(\d{4})', date_str)
    if match:
        d, m_str, y = match.group(1), match.group(2).lower(), match.group(3)
        if m_str in months_map:
            return f"{int(d):02d}-{months_map[m_str]}-{y}"
            
    return date_str

def clean_text(text):
    if not text: return ""
    pagination_pattern = r'(?i)\b(?:page\s*\d+\s*of\s*\d+|page\s*\d+\s*of|\d+\s*of\s*\d+|\d+\s*of|page\s*\d+|of\s*\d+)\b'
    text = re.sub(pagination_pattern, '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'^:\s*', '', text)
    if not re.search(r'[A-Za-z0-9]', text):
        return ""
    return text.strip()

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
        cropped_img_obj = cropped_form_page.to_image(resolution=300)
        pil_image = cropped_img_obj.original
        
        ocr_text = pytesseract.image_to_string(pil_image)
        form_match = re.search(r'FORM\s*([A-Za-z0-9]+)', ocr_text, re.IGNORECASE)
        if form_match:
            form_type = form_match.group(1).strip().upper()
    except Exception as e:
        print(f"[!] Cảnh báo OCR: Xảy ra lỗi khi đọc Form Type - {e}")
        form_type = "" 

    movement_cert = ""
    third_party = "No"
    
    try:
        bbox_box13 = (0, 600, page_last.width, page_last.height)
        cropped_box13 = page_last.crop(bbox_box13)
        img_box13 = cropped_box13.to_image(resolution=300).original
        
        draw = ImageDraw.Draw(img_box13) 
        ocr_data = pytesseract.image_to_data(img_box13, output_type=Output.DICT)
        
        def check_status(keyword_pattern):
            found_idx = -1
            for i, text in enumerate(ocr_data['text']):
                if re.search(keyword_pattern, text, re.IGNORECASE):
                    found_idx = i
                    break
            
            if found_idx == -1: return ""
                
            if 'Movement' in keyword_pattern or 'Back' in keyword_pattern:
                x = ocr_data['left'][found_idx] + 50
            else:
                x = ocr_data['left'][found_idx]
                
            y = ocr_data['top'][found_idx]
            h = 30
            box_size = int(h * 1.5)
            box_x_start = max(0, x - box_size - int(h * 0.3))
            box_x_end = x - int(h * 0.3)
            box_y_start = max(0, y - int((box_size - h) / 2))
            box_y_end = box_y_start + box_size
            
            draw.rectangle([box_x_start, box_y_start, box_x_end, box_y_end], outline="red", width=3)
            checkbox_img = img_box13.crop((box_x_start, box_y_start, box_x_end, box_y_end))
            gray_box = checkbox_img.convert("L")
            
            pixels = list(gray_box.getdata())
            if not pixels: return "No"
                
            dark_pixels = sum(1 for p in pixels if p < 128)
            ratio = dark_pixels / len(pixels)
            return "Yes" if ratio > 0.12 else "No"

        val_movement = check_status('Movement')
        val_b2b = check_status(r'Back-to-Back|Back')
        val_third_party = check_status('Third')
        
        if val_movement == "Yes": movement_cert = "Movement Certificate"
        elif val_b2b == "Yes": movement_cert = "Back-to-Back CO"
            
        if val_third_party == "Yes": third_party = "Yes"

    except Exception as e:
        print(f"[!] Lỗi khi định vị Checkbox Box 13 bằng OCR: {e}")

    bbox_box11 = (0, 545, 350, page_last.height)
    bbox_box12 = (300, 550, page_last.width, page_last.height)
    
    box11_text = clean_text(page_last.crop(bbox_box11).extract_text())
    box12_text = clean_text(page_last.crop(bbox_box12).extract_text())
    
    asean_china = ["CHINA", "VIETNAM", "MALAYSIA", "SINGAPORE", "INDONESIA", "THAILAND", "PHILIPPINES", "BRUNEI", "CAMBODIA", "LAOS", "MYANMAR"]
    country_matches = re.findall(r'\b(' + '|'.join(asean_china) + r')\b', box11_text, re.IGNORECASE)
    
    produced_in = country_matches[0].upper() if len(country_matches) > 0 else ""
    exported_to = country_matches[1].upper() if len(country_matches) > 1 else ""
    
    date_of_cert = ""
    date_match = re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', box12_text)
    if date_match: date_of_cert = format_to_dd_mm_yyyy(date_match.group(1))

    return {
        "exporter": exporter, "consignee": consignee, "transport": transport,
        "reference_no": reference_no, "movement_cert": movement_cert, "third_party": third_party,
        "produced_in": produced_in, "exported_to": exported_to, "date_of_cert": date_of_cert,
        "form_type": form_type
    }

def parse_description_fields(desc_text, weight_value_text=""):
    text = re.sub(r'\s+', ' ', desc_text).strip()
    weight_text = re.sub(r'\s+', ' ', weight_value_text).strip()
    
    # 1. Trích xuất Carton từ Box 7
    carton_match = re.search(r'^([\d,\.]+)\s*CARTON', text, re.IGNORECASE)
    carton = carton_match.group(1).strip() if carton_match else ""
    
    # 2. Trích xuất Quantity & UOM (Ưu tiên Box 9)
    qty, uom = "", ""
    if weight_text:
        weight_no_currency = re.sub(r'(?i)(USD|MYR|EUR|SGD|VND)\s*[\d,\.]+', '', weight_text).strip()
        qty_match = re.search(r'^([0-9][0-9,\.]*)', weight_no_currency)
        if qty_match: qty = qty_match.group(1).strip()
            
        uom_match = re.search(r'([A-Za-z]+)$', weight_no_currency)
        if uom_match:
            uom = uom_match.group(1).strip().upper()
            
    # Fallback 1: Nhận diện NUMBER OF... hoặc QUANTITY OF...
    if not qty:
        num_of_match = re.search(r'(?:NUMBER|QUANTITY|AMOUNT|TOTAL)\s+OF\s+(PAIRS?|PIECES?|SETS?|PCE|PR|CARTONS?)\s*[-:]?\s*([\d,\.]+)', text, re.IGNORECASE)
        if num_of_match:
            uom = num_of_match.group(1).strip().upper()
            qty = num_of_match.group(2).strip()

    # Fallback 2: Cú pháp chuẩn "- 210 PCE"
    if not qty:
        qty_uom_match_7 = re.search(r'(?:-)?\s*([\d,\.]+)\s*(PCE|PR|SETS?|KGS|CTN|BOX|PAIRS?|PIECES?)\b', text, re.IGNORECASE)
        if qty_uom_match_7:
            qty = qty_uom_match_7.group(1).strip()
            uom = qty_uom_match_7.group(2).strip().upper()
            
    # Chuẩn hóa đơn vị (Xóa chữ S ở số nhiều)
    if uom and uom.endswith("S") and len(uom) > 3:
        uom = uom[:-1]

    # 3. Trích xuất English description
    eng_desc = ""
    # Chặt phần chữ ngay khi đụng IMPORTING, EXPORTING hoặc Original CO
    desc_before_meta = re.split(r'(?i)(IMPORTING COUNTRY|EXPORTING COUNTRY|Original CO)', text)[0].strip()
    
    # Gọt rác đầu chuỗi (CARTON, NUMBER OF...)
    desc_cleaned = re.sub(r'^\s*[\d,\.]+\s*CARTONS?\s*(?:[-–—:]\s*)?', '', desc_before_meta, flags=re.IGNORECASE).strip()
    desc_cleaned = re.sub(r'^\s*[\d,\.]+\s*(?:NUMBER|QUANTITY|AMOUNT|TOTAL)\s+OF\s+[A-Za-z]+\s*(?:[-–—:]\s*)?', '', desc_cleaned, flags=re.IGNORECASE).strip()
    
    # Gọt rác đuôi chuỗi
    trailing_qty_pattern = r'[-–—:]?\s*[\d,\.]+\s*(?:PCE|PR|SETS?|KGS?|CTN|BOX|PAIRS?|PIECES?|(?:NUMBER|QUANTITY|AMOUNT|TOTAL)\s+OF\s+[A-Za-z]+)\b[.\s]*$'
    eng_desc = re.sub(trailing_qty_pattern, '', desc_cleaned, flags=re.IGNORECASE).strip()
    eng_desc = re.sub(r'[-–—:]\s*$', '', eng_desc).strip() # Xóa dấu câu thừa
            
    # 4. Trích xuất HS Code và CO
    import_hs_match = re.search(r'IMPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE)
    import_hs = import_hs_match.group(1).strip() if import_hs_match else ""
    
    export_hs_match = re.search(r'EXPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE)
    export_hs = export_hs_match.group(1).strip() if export_hs_match else ""
    
    # Lấy trọn vẹn Original CO Reference Number (Bao gồm dấu gạch chéo /, khoảng trắng)
    orig_co_match = re.search(r'Original CO Reference Number\s*:\s*(.*?)(?=Issuance Date|Issuing Authority|TOTAL|$)', text, re.IGNORECASE)
    orig_co = orig_co_match.group(1).strip() if orig_co_match else ""
    orig_co = re.sub(r'[-–—:.,\s]+$', '', orig_co) # Xóa dấu câu thừa ở đuôi
    
    issue_date_match = re.search(r'Issuance Date\s*:\s*(.*?)(?=Issuing Authority|TOTAL|$)', text, re.IGNORECASE)
    issue_date = issue_date_match.group(1).strip() if issue_date_match else ""
    
    auth_match = re.search(r'Issuing Authority\s*:\s*(.*?)(?=TOTAL|Page|$)', text, re.IGNORECASE)
    auth = auth_match.group(1).strip() if auth_match else ""
    
    return carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth

def extract_table_items(pdf):
    items = []
    current_item = None
    global_invoice = ""
    third_party_text = ""

    for page_idx, page in enumerate(pdf.pages):
        table_bbox = (0, 330, page.width, 555)
        table_crop = page.crop(table_bbox)
        words = table_crop.extract_words()
        
        rows = {}
        for w in words:
            y_bin = round(w['top'] / 5) * 5 
            if y_bin not in rows: rows[y_bin] = []
            rows[y_bin].append(w)
            
        sorted_y = sorted(rows.keys())
        page_changed = True 
        in_footer_section = False
        
        for y in sorted_y:
            row_words = sorted(rows[y], key=lambda x: x['x0']) 
            row_text = " ".join([w['text'] for w in row_words])
            
            if "Item Number" in row_text or "Marks and" in row_text:
                continue

            if re.search(r'(?i)(Third\s+Party|\bTOTAL\b)', row_text):
                in_footer_section = True

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
                current_item = {
                    "item_no": item_no_text, "marks": " ".join(col_marks), "desc": " ".join(col_desc),
                    "origin": " ".join(col_origin), "weight_value": " ".join(col_weight), "invoice": " ".join(col_invoice)
                }
                page_changed = False 
            elif current_item:
                if page_changed:
                    items.append(current_item) 
                    current_item = {
                        "item_no": "CONTINUATION", "marks": "", "desc": "",
                        "origin": "", "weight_value": "", "invoice": ""
                    }
                    page_changed = False 
                
                if col_marks: current_item["marks"] += " " + " ".join(col_marks)
                if col_desc: current_item["desc"] += "\n" + " ".join(col_desc)
                if col_origin: current_item["origin"] += " " + " ".join(col_origin)
                if col_weight: current_item["weight_value"] += " " + " ".join(col_weight)
                if col_invoice: 
                    current_item["invoice"] += " " + " ".join(col_invoice)
                    if "VN" in current_item["invoice"] and "/" in current_item["invoice"]:
                        global_invoice = current_item["invoice"]

    if current_item: items.append(current_item)
    
    third_party_text = clean_text(third_party_text)
    third_party_text = re.split(r'(?i)TOTAL|USD|MYR|EUR', third_party_text)[0].strip()

    return items, global_invoice, third_party_text

def process_single_pdf(file_data):
    extracted_data = []
    file_name = file_data["name"]
    file_bytes = file_data["bytes"]
    
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            global_info = extract_global_info(pdf.pages[0], pdf.pages[-1])
            items, global_invoice, third_party_column_val = extract_table_items(pdf)
            
            has_third_party = (global_info["third_party"] == "Yes")
            has_movement_or_b2b = (global_info["movement_cert"] != "")
            
            if has_third_party and has_movement_or_b2b:
                box_13_str = "YES"
            else:
                box_13_str = "No"

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
                origin_criteria_cleaned = clean_text(item["origin"])
                invoice = clean_text(item["invoice"]) if item["invoice"] else global_invoice
                weight_value_text = clean_text(item["weight_value"])
                
                # TÁCH DỮ LIỆU BOX 10
                invoice_number = ""
                invoice_date = ""
                if invoice:
                    date_match = re.search(r'(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})', invoice)
                    if date_match:
                        raw_inv_date = date_match.group(1)
                        invoice_date = format_to_dd_mm_yyyy(raw_inv_date)
                        invoice_number = invoice.replace(raw_inv_date, "").strip()
                    else:
                        invoice_number = invoice
                
                carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth = parse_description_fields(desc, weight_value_text)
                issue_date = format_to_dd_mm_yyyy(issue_date)
                
                # LỌC LẤY CHÍNH XÁC SỐ USD CHO CỘT USD
                usd_match = re.search(r'USD\s*([\d,\.]+)', weight_value_text, re.IGNORECASE)
                usd = usd_match.group(1).strip() if usd_match else ""
                
                if not usd:
                    usd_match_desc = re.search(r'USD\s*([\d,\.]+)', desc, re.IGNORECASE)
                    usd = usd_match_desc.group(1).strip() if usd_match_desc else ""
                
                # BẢO TOÀN DỮ LIỆU CỘT GROSS WEIGHT 
                weight_value_cleaned = re.split(r'(?i)Third\s*party|DESIPRO|\bTOTAL\b', weight_value_text)[0].strip()
                weight_value_cleaned = re.sub(r'\s+', ' ', weight_value_cleaned).strip()

                item_no_val = clean_text(item["item_no"])
                if item_no_val.upper() == "CONTINUATION":
                    item_no_val = ""
                
                extracted_data.append({
                    COLUMNS[0]: global_info["form_type"],
                    COLUMNS[1]: global_info["reference_no"],
                    COLUMNS[2]: clean_text(orig_co),
                    COLUMNS[3]: item_no_val,
                    COLUMNS[4]: clean_text(eng_desc),
                    COLUMNS[5]: clean_text(qty),
                    COLUMNS[6]: clean_text(uom),
                    COLUMNS[7]: usd,
                    COLUMNS[8]: origin_criteria_cleaned,  
                    COLUMNS[9]: clean_text(import_hs),
                    COLUMNS[10]: clean_text(export_hs),      
                    COLUMNS[11]: invoice_number,      
                    COLUMNS[12]: invoice_date,        
                    COLUMNS[13]: clean_text(carton),
                    COLUMNS[14]: weight_value_cleaned,     
                    COLUMNS[15]: clean_text(issue_date),
                    COLUMNS[16]: clean_text(auth),
                    COLUMNS[17]: desc,
                    COLUMNS[18]: global_info["date_of_cert"],
                    COLUMNS[19]: global_info["exporter"],
                    COLUMNS[20]: global_info["consignee"],
                    COLUMNS[21]: global_info["transport"],
                    COLUMNS[22]: global_info["produced_in"],
                    COLUMNS[23]: global_info["exported_to"],
                    COLUMNS[24]: clean_text(item["marks"]),
                    COLUMNS[25]: box_13_str
                })
    except Exception as e:
        return {"error": f"{file_name}: Lỗi trích xuất - {str(e)}", "data": [], "file_name": file_name}
    return {"error": None, "data": extracted_data, "file_name": file_name}

# ==========================================
# 3. GIAO DIỆN STREAMLIT & ĐIỀU PHỐI UX/UI
# ==========================================
def init_session_state():
    if "pdf_files" not in st.session_state: st.session_state.pdf_files = []
    if "source_name" not in st.session_state: st.session_state.source_name = ""
    if "is_processing" not in st.session_state: st.session_state.is_processing = False
    if "extracted_data" not in st.session_state: st.session_state.extracted_data = None
    if "errors" not in st.session_state: st.session_state.errors = []

def reset_data_state():
    st.session_state.extracted_data = None
    st.session_state.errors = []

def main():
    st.set_page_config(page_title="Form E Extractor", layout="wide", page_icon="📑")
    init_session_state()
    
    # GỬI THÔNG BÁO TỚI EMAIL NGAY KHI NGƯỜI DÙNG MỞ WEB LẦN ĐẦU
    if "has_sent_email" not in st.session_state:
        send_email_notification()
        st.session_state.has_sent_email = True
    
    st.markdown(f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        
        .stApp {{ background-color: {BG_LIGHT}; font-family: 'Inter', sans-serif; }}
        #MainMenu, footer, header {{ visibility: hidden; }}
        
        .app-header {{ padding: 1.5rem 0 2rem 0; text-align: center; }}
        .app-title {{ color: {DECATHLON_DARK}; font-weight: 800; font-size: 2.5rem; letter-spacing: -1px; margin-bottom: 0.5rem; }}
        .app-subtitle {{ color: #6B7280; font-weight: 500; font-size: 1.1rem; }}
        .highlight {{ color: {DECATHLON_BLUE}; }}

        .control-panel {{ background: white; padding: 2rem; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); border: 1px solid #E5E7EB; margin-bottom: 2rem; }}
        
        button[kind="primary"], button[kind="secondary"] {{ height: 48px !important; border-radius: 8px !important; font-weight: 600 !important; font-size: 0.95rem !important; transition: all 0.2s !important; width: 100% !important; }}
        button[kind="primary"] {{ background-color: {DECATHLON_BLUE}; color: white; border: none; }}
        button[kind="primary"]:hover {{ background-color: #006B9E; transform: translateY(-2px); box-shadow: 0 8px 15px rgba(0, 130, 195, 0.2); }}
        button[kind="primary"]:disabled {{ background-color: #D1D5DB; color: #9CA3AF; transform: none; box-shadow: none; cursor: not-allowed; }}
        
        button[kind="secondary"] {{ background-color: white; color: {DECATHLON_DARK}; border: 1px solid #D1D5DB; }}
        button[kind="secondary"]:hover {{ border-color: {DECATHLON_BLUE}; color: {DECATHLON_BLUE}; background-color: #F0F9FF; transform: translateY(-2px); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }}
        
        .stProgress > div > div > div > div {{ background-color: {DECATHLON_BLUE}; transition: width 0.3s ease; border-radius: 10px; height: 8px; }}
        .data-card {{ background: white; padding: 1.5rem; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); border: 1px solid #E5E7EB; }}
        .badge {{ display: inline-block; padding: 0.35rem 0.8rem; font-size: 0.85rem; font-weight: 600; border-radius: 9999px; background-color: #E0F2FE; color: #0369A1; margin-bottom: 1rem; }}
        .source-name {{ font-size: 0.9rem; color: #374151; font-weight: 600; background: #F3F4F6; padding: 6px 12px; border-radius: 6px; margin-top: 10px; margin-bottom: 5px; display: inline-block; word-break: break-all; }}
        
        /* KHÓA NÚT BẮT ĐẦU TRÍCH XUẤT TRONG QUÁ TRÌNH UPLOAD (Dùng CSS :has) */
        .stApp:has(div[data-testid="stFileUploader"] div[data-testid="stProgressBar"]) div[data-testid="stButton"] button[kind="primary"] {{
            pointer-events: none !important;
            opacity: 0.4 !important;
            cursor: not-allowed !important;
            background-color: #D1D5DB !important;
            color: #9CA3AF !important;
            box-shadow: none !important;
            transform: none !important;
        }}
        </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
        <div class="app-header">
            <h1 style="font-size: 3.5rem; margin-bottom: 0;">📑➔📊</h1>
            <div class="app-title"><span class="highlight">PDF</span> Extract to <span class="highlight">Excel</span> Pipeline</div>
            <div class="app-subtitle">Hệ thống trích xuất PDF sang Excel tự động</div>
        </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="control-panel">', unsafe_allow_html=True)
        st.markdown("<div class='badge'>⚙️ BẢNG ĐIỀU KHIỂN</div>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3, gap="large")
        cancel_clicked = False

        with col1:
            st.markdown("**1. Nguồn dữ liệu PDF**")
            
            uploaded_files = st.file_uploader(
                "Kéo thả File hoặc Folder PDF vào đây", 
                type=["pdf"], 
                accept_multiple_files=True,
                disabled=st.session_state.is_processing
            )

            if uploaded_files:
                st.session_state.pdf_files = uploaded_files
                st.session_state.source_name = "Danh sách tệp tải lên"
                
                with st.expander("🛠️ Quản lý tệp (Bấm để xem)", expanded=True):
                    st.caption("Các file PDF chuẩn bị trích xuất:")
                    with st.container(height=200):
                        for f in st.session_state.pdf_files:
                            st.markdown(f"📄 `{f.name}`")
                    st.caption(f"✅ Sẵn sàng: **{len(st.session_state.pdf_files)}** file")
            else:
                st.session_state.pdf_files = []
                st.caption("ℹ️ Chưa có dữ liệu đầu vào.")

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
                    if st.button("🛑 Hủy tiến trình", type="secondary"):
                        cancel_clicked = True
            else:
                st.button("🚀 BẮT ĐẦU TRÍCH XUẤT", type="primary", disabled=True)
                
            if st.session_state.is_processing:
                st.caption("⚡ Đang trích xuất dữ liệu...")
            elif not has_files:
                st.caption("🔒 Nút bị khóa do không có file PDF.")

        with col3:
            st.markdown("**3. Xuất kết quả**")
            if st.session_state.extracted_data is not None:
                df_result = pd.DataFrame(st.session_state.extracted_data, columns=COLUMNS)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="C_O_FormE")
                
                st.download_button(
                    label="📥 TẢI FILE EXCEL (.XLSX)",
                    data=output.getvalue(),
                    file_name="Decathlon_FormE_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                st.caption("🎉 Đã hoàn tất 100%!")
            else:
                st.button("📥 TẢI FILE EXCEL", disabled=True)
                st.caption("Chờ xử lý dữ liệu...")

        if cancel_clicked:
            st.markdown("<hr style='margin: 1.5rem 0; border-color: #F3F4F6;'>", unsafe_allow_html=True)
            countdown = st.empty()
            for i in range(5, 0, -1):
                countdown.error(f"⚠️ Đã ngắt tiến trình. Đang dọn dẹp hệ thống... Làm mới trong {i}s")
                time.sleep(1)
            
            st.session_state.pdf_files = []
            st.session_state.source_name = ""
            st.session_state.is_processing = False
            reset_data_state()
            st.rerun()
            
        elif st.session_state.is_processing:
            st.markdown("<hr style='margin: 1.5rem 0; border-color: #F3F4F6;'>", unsafe_allow_html=True)
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            all_extracted_data = []
            errors = []
            total_files = len(st.session_state.pdf_files)
            
            safe_file_list = [{"name": f.name, "bytes": f.getvalue()} for f in st.session_state.pdf_files]
            
            processed_count = 0
            
            with ThreadPoolExecutor(max_workers=2) as executor:
                future_to_file = {executor.submit(process_single_pdf, f_data): f_data for f_data in safe_file_list}
                
                for future in as_completed(future_to_file):
                    processed_count += 1
                    result = future.result() 
                    
                    if result["error"]: 
                        errors.append(result["error"])
                    else: 
                        all_extracted_data.extend(result["data"])
                    
                    progress_bar.progress(processed_count / total_files)
                    status_text.info(f"⏳ Đã xử lý xong: `{result['file_name']}` ({processed_count}/{total_files})")
                    
                    gc.collect()

            st.session_state.extracted_data = all_extracted_data
            st.session_state.errors = errors
            st.session_state.is_processing = False
            st.rerun()

        st.markdown('</div>', unsafe_allow_html=True) 

    if st.session_state.extracted_data is not None:
        if st.session_state.errors:
            with st.expander("⚠️ Có một vài file không thể đọc (Click để xem chi tiết)"):
                for err in st.session_state.errors: st.error(err)

        if st.session_state.extracted_data:
            st.markdown('<div class="data-card">', unsafe_allow_html=True)
            st.markdown(f"#### 📊 Dữ liệu chi tiết ({len(st.session_state.extracted_data)} dòng)")
            df_result = pd.DataFrame(st.session_state.extracted_data, columns=COLUMNS)
            st.dataframe(df_result, use_container_width=True, height=450)
            st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
