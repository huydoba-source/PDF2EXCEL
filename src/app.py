import streamlit as st
import pandas as pd
import os
import glob
import io
import re
import time
import pdfplumber
import gc
from concurrent.futures import ProcessPoolExecutor, as_completed
import tkinter as tk
from tkinter import filedialog

# --- THƯ VIỆN BỔ SUNG CHO OCR ---
import pytesseract
from pytesseract import Output
from PIL import Image, ImageDraw

# [TÙY CHỌN] CẤU HÌNH ĐƯỜNG DẪN TESSERACT BẮT BUỘC (Dành cho môi trường Local Windows)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ==========================================
# 1. CẤU HÌNH CỘT DỮ LIỆU ĐẦU RA (EXCEL & AUDIO SYNC)
# ==========================================
COLUMNS = [
    "Products consigned from (Exporter's business name, address, country)",
    "Products consigned to (Consignee's name, address, country)",
    "Means of transport and route (as far as known)",
    "Reference No",
    "Item Number",
    "Marks and numbers on packages",
    "Number and type of packages, description of products",
    "Origin criteria (see Overleaf Notes)",
    "Gross weight or net weight or other quantity, and value",
    "Invoice Number",          # --- TÁCH TỪ BOX 10 ---
    "Date of invoices",        # --- TÁCH TỪ BOX 10 ---
    "CARTON",
    "English description",     # --- CỘT MỚI TỪ AUDIO ---
    "IMPORTING COUNTRY HS CODE",
    "EXPORTING COUNTRY HS CODE",
    "Original CO Reference Number",
    "Issuance Date",
    "Issuing Authority",
    "Quantity",                # --- CỘT MỚI (Thay thế PCE) ---
    "UOM",                     # --- CỘT MỚI (Thay thế PCE) ---
    "produced in",
    "exported to",
    "Date of certification",
    "Form", 
    "USD",                     # --- CHỈ LẤY USD ---
    "Box 13",                  # --- CỘT GỘP TỪ AUDIO ---
    "Third Party"              # --- CHUYỂN TỪ ROW SANG COLUMN ---
]

DECATHLON_BLUE = "#0082C3"
DECATHLON_DARK = "#1F2937"
BG_LIGHT = "#F9FAFB"

# ==========================================
# 2. CORE FUNCTIONS & DATA CLEANING
# ==========================================
def clean_text(text):
    if not text: return ""
    pagination_pattern = r'(?i)\b(?:page\s*\d+\s*of\s*\d+|page\s*\d+\s*of|\d+\s*of\s*\d+|\d+\s*of|page\s*\d+|of\s*\d+)\b'
    text = re.sub(pagination_pattern, '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'^:\s*', '', text)
    if not re.search(r'[A-Za-z0-9]', text): return ""
    return text.strip()

def standardize_date(date_str):
    if not date_str: return ""
    date_str = re.sub(r'(?i)(Page.*|TOTAL.*)', '', date_str).strip()
    try:
        # Sử dụng Pandas để ép kiểu và format về chuẩn DD-MMM-YYYY (VD: 19-APR-2026)
        dt = pd.to_datetime(date_str, format='mixed', dayfirst=True)
        if pd.notnull(dt):
            return dt.strftime("%d-%b-%Y").upper()
    except:
        pass
    return date_str

def split_invoice(invoice_text):
    if not invoice_text: return "", ""
    # Tìm chuỗi ngày tháng (VD: 19/04/2026, 01-APR-2026)
    date_pattern = r'(\d{1,2}[\/\-\s]+[A-Za-z]{3,}[\/\-\s]+\d{2,4}|\d{1,2}[\/\-\s]+\d{1,2}[\/\-\s]+\d{2,4})'
    match = re.search(date_pattern, invoice_text)
    if match:
        date_part = match.group(1)
        inv_part = invoice_text[:match.start()].strip()
        # Dọn dẹp rác (CN/VN) trôi nổi ở phần đuôi mã hóa đơn
        inv_part = re.sub(r'(CN|VN)\s*$', '', inv_part, flags=re.IGNORECASE).strip()
        return inv_part, standardize_date(date_part)
    return invoice_text, ""

def extract_global_info(page_first, page_last):
    # --- LẤY BOX 1, 2, 3, 4 (PAGE FIRST) ---
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

    # --- OCR LOẠI FORM (PAGE FIRST) ---
    form_type = ""
    try:
        bbox_form = (290, 0, page_first.width, 150)
        pil_image = page_first.crop(bbox_form).to_image(resolution=300).original
        ocr_text = pytesseract.image_to_string(pil_image)
        form_match = re.search(r'FORM\s*([A-Za-z0-9]+)', ocr_text, re.IGNORECASE)
        if form_match: form_type = form_match.group(1).strip().upper()
    except: pass 

    # --- OCR BOX 13 CHECKBOX (PAGE LAST) ---
    movement_cert = "No"
    third_party_cert = "No"
    try:
        bbox_box13 = (0, 600, page_last.width, page_last.height)
        img_box13 = page_last.crop(bbox_box13).to_image(resolution=300).original
        ocr_data = pytesseract.image_to_data(img_box13, output_type=Output.DICT)
        
        def check_status(keyword_pattern):
            found_idx = -1
            for i, text in enumerate(ocr_data['text']):
                if re.search(keyword_pattern, text, re.IGNORECASE):
                    found_idx = i
                    break
            if found_idx == -1: return ""
            x = ocr_data['left'][found_idx] + 50 if keyword_pattern == 'Movement' else ocr_data['left'][found_idx]
            y = ocr_data['top'][found_idx]
            h = 30
            box_size = int(h * 1.5)
            box_x_start = max(0, x - box_size - int(h * 0.3))
            box_x_end = x - int(h * 0.3)
            box_y_start = max(0, y - int((box_size - h) / 2))
            
            gray_box = img_box13.crop((box_x_start, box_y_start, box_x_end, box_y_start + box_size)).convert("L")
            pixels = list(gray_box.getdata())
            if not pixels: return "No"
            ratio = sum(1 for p in pixels if p < 128) / len(pixels)
            return "Yes" if ratio > 0.12 else "No"

        val_movement = check_status('Movement')
        if val_movement: movement_cert = val_movement
        val_third_party = check_status('Third')
        if val_third_party: third_party_cert = val_third_party

    except: pass

    # --- LẤY BOX 11 & 12 (PAGE LAST) ---
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
    if date_match: date_of_cert = date_match.group(1)

    return {
        "exporter": exporter, "consignee": consignee, "transport": transport,
        "reference_no": reference_no, "movement_cert": movement_cert, "third_party_cert": third_party_cert,
        "produced_in": produced_in, "exported_to": exported_to, "date_of_cert": date_of_cert,
        "form_type": form_type
    }

def parse_description_fields(desc_text):
    text = re.sub(r'\s+', ' ', desc_text).strip()
    
    # Bóc tách CARTON
    carton_match = re.search(r'([\d,\.]+)\s*CARTON', text, re.IGNORECASE)
    carton = carton_match.group(1).strip() if carton_match else ""
    
    # Tách vùng mô tả Sản phẩm (English Description) và Quantity + UOM
    qty = ""
    uom = ""
    english_desc = ""
    
    prod_area_match = re.search(r'(.*?)(?=IMPORTING COUNTRY|$)', text, re.IGNORECASE)
    prod_area = prod_area_match.group(1).strip() if prod_area_match else text
    
    # Tìm Số lượng và Đơn vị nằm sát nhau ở cuối đoạn (VD: - 18 PCE, 2 PIECE, 10 PR)
    qty_uom_match = re.search(r'[-\s]*([\d,\.]+)\s+([A-Za-z]+)$', prod_area, re.IGNORECASE)
    if qty_uom_match:
        qty = qty_uom_match.group(1)
        uom = qty_uom_match.group(2).upper()
        desc_part = prod_area[:qty_uom_match.start()].strip()
    else:
        # Fallback nếu Regex 1 trượt
        qty_match = re.search(r'([\d,\.]+)\s*(PCE|PIECE|PR|PAIRS|SETS|UNITS)', prod_area, re.IGNORECASE)
        if qty_match:
            qty = qty_match.group(1)
            uom = qty_match.group(2).upper()
            desc_part = prod_area[:qty_match.start()].strip()
        else:
            desc_part = prod_area
            
    # Gọt dũa mô tả Tiếng Anh (Xóa số Carton và ký tự thừa ở đầu)
    desc_part = re.sub(r'[\d,\.]+\s*CARTON', '', desc_part, flags=re.IGNORECASE).strip()
    english_desc = re.sub(r'^[-\:]\s*', '', desc_part).strip()

    # Bóc tách HS CODE & CO
    import_hs_match = re.search(r'IMPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE)
    import_hs = import_hs_match.group(1).strip() if import_hs_match else ""
    
    export_hs_match = re.search(r'EXPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE)
    export_hs = export_hs_match.group(1).strip() if export_hs_match else ""
    
    orig_co_match = re.search(r'Original CO Reference Number\s*:\s*([A-Za-z0-9\-]+)', text, re.IGNORECASE)
    orig_co = orig_co_match.group(1).strip() if orig_co_match else ""
    
    issue_date_match = re.search(r'Issuance Date\s*:\s*(.*?)(?=Issuing Authority|TOTAL|$)', text, re.IGNORECASE)
    issue_date = issue_date_match.group(1).strip() if issue_date_match else ""
    
    auth_match = re.search(r'Issuing Authority\s*:\s*(.*?)(?=TOTAL|Page|$)', text, re.IGNORECASE)
    auth = auth_match.group(1).strip() if auth_match else ""
    
    return carton, qty, uom, english_desc, import_hs, export_hs, orig_co, issue_date, auth

def extract_table_items(pdf):
    items = []
    current_item = None
    global_invoice = ""

    for page_idx, page in enumerate(pdf.pages):
        table_bbox = (0, 330, page.width, 555)
        words = page.crop(table_bbox).extract_words()
        
        rows = {}
        for w in words:
            y_bin = round(w['top'] / 5) * 5 
            if y_bin not in rows: rows[y_bin] = []
            rows[y_bin].append(w)
            
        sorted_y = sorted(rows.keys())
        page_changed = True 
        
        for y in sorted_y:
            row_words = sorted(rows[y], key=lambda x: x['x0']) 
            row_text = " ".join([w['text'] for w in row_words])
            
            if "Item Number" in row_text or "Marks and" in row_text:
                continue

            col_item_no = [w['text'] for w in row_words if 35 <= w['x0'] < 77]
            col_marks   = [w['text'] for w in row_words if 77 <= w['x0'] < 150]
            col_desc    = [w['text'] for w in row_words if 150 <= w['x0'] < 310]
            col_origin  = [w['text'] for w in row_words if 310 <= w['x0'] < 398]
            col_weight  = [w['text'] for w in row_words if 398 <= w['x0'] < 480]
            col_invoice = [w['text'] for w in row_words if w['x0'] >= 480]

            item_no_text = " ".join(col_item_no).strip()
            check_number = item_no_text.replace(".", "").replace(",", "").strip()

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
                    current_item = {"item_no": "CONTINUATION", "marks": "", "desc": "", "origin": "", "weight_value": "", "invoice": ""}
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
    return items, global_invoice

def process_single_pdf(file_path):
    extracted_data = []
    file_name = os.path.basename(file_path)
    try:
        with pdfplumber.open(file_path) as pdf:
            page_first = pdf.pages[0]
            page_last = pdf.pages[-1]
            global_info = extract_global_info(page_first, page_last)
            items, global_invoice = extract_table_items(pdf)
            
            if items:
                all_desc = " \n ".join([it["desc"] for it in items])
                all_weight_value = " \n ".join([it["weight_value"] for it in items])

                # --- 1. LẤY THIRD PARTY CHO CỘT (LOẠI BỎ TOTAL/USD) ---
                tp_match = re.search(r'(?i)(Third[-\s]*Party.*?(?=TOTAL|USD|Page|$))', all_desc)
                global_third_party = tp_match.group(1).strip() if tp_match else ""
                global_third_party = re.sub(r'(?i)TOTAL.*', '', global_third_party).strip()

                # --- 2. CHỈ LẤY ĐÚNG MỆNH GIÁ USD ---
                usd_match = re.search(r'USD\s*([\d,\.]+)', all_weight_value + " " + all_desc, re.IGNORECASE)
                global_usd = usd_match.group(1).strip() if usd_match else ""

                # --- 3. LÀM SẠCH VÀ GỘP DÒNG ---
                merged_items = []
                for item in items:
                    # Xóa rác Third Party & Total khỏi Description để nó không bị dính vào Hàng hóa
                    item["desc"] = re.sub(r'(?i)(Third[-\s]*Party.*?(?=TOTAL|USD|Page|$))', '', item["desc"]).strip()
                    item["desc"] = re.sub(r'(?i)TOTAL\s*:?\s*[\d,\.]+\s*[A-Za-z]*', '', item["desc"]).strip()
                    item["desc"] = re.sub(r'(?i)(USD|MYR|EUR)\s*[\d,\.]+', '', item["desc"]).strip()
                    item["desc"] = clean_text(item["desc"])

                    if item["item_no"] == "CONTINUATION":
                        if merged_items: 
                            merged_items[-1]["marks"] = (merged_items[-1]["marks"] + "\n" + item["marks"]).strip()
                            merged_items[-1]["desc"] = (merged_items[-1]["desc"] + "\n" + item["desc"]).strip()
                            merged_items[-1]["origin"] = (merged_items[-1]["origin"] + " " + item["origin"]).strip()
                            merged_items[-1]["weight_value"] = (merged_items[-1]["weight_value"] + "\n" + item["weight_value"]).strip()
                            merged_items[-1]["invoice"] = (merged_items[-1]["invoice"] + "\n" + item["invoice"]).strip()
                    else:
                        merged_items.append(item)
                
                # Bỏ đi những dòng trống (VD dòng Total cũ đã bị xóa sạch chữ)
                final_items = [it for it in merged_items if re.search(r'[A-Za-z0-9]', it["desc"])]
            else:
                final_items = []
            
            for item in final_items:
                desc = clean_text(item["desc"])
                invoice_raw = clean_text(item["invoice"]) if item["invoice"] else global_invoice
                
                # Cắt Invoice Number & Date
                inv_number, inv_date = split_invoice(invoice_raw)
                
                # Bóc Box 7 (Có Quantity, UOM, English Desc)
                carton, qty, uom, english_desc, import_hs, export_hs, orig_co, issue_date, auth = parse_description_fields(desc)
                
                # Format Dates
                issue_date = standardize_date(issue_date)
                date_of_cert = standardize_date(global_info["date_of_cert"])
                
                # Xây dựng Box 13 Gộp
                box13_arr = []
                if global_info["movement_cert"] == "Yes": box13_arr.append("Movement Certificate")
                if global_info["third_party_cert"] == "Yes": box13_arr.append("Third Party Invoicing")
                box13_val = ", ".join(box13_arr)
                
                extracted_data.append({
                    COLUMNS[0]: global_info["exporter"],
                    COLUMNS[1]: global_info["consignee"],
                    COLUMNS[2]: global_info["transport"],
                    COLUMNS[3]: global_info["reference_no"],
                    COLUMNS[4]: clean_text(item["item_no"]),
                    COLUMNS[5]: clean_text(item["marks"]),
                    COLUMNS[6]: desc,
                    COLUMNS[7]: clean_text(item["origin"]),
                    COLUMNS[8]: clean_text(item["weight_value"]),
                    COLUMNS[9]: inv_number,          
                    COLUMNS[10]: inv_date,            
                    COLUMNS[11]: clean_text(carton),
                    COLUMNS[12]: clean_text(english_desc), 
                    COLUMNS[13]: clean_text(import_hs),
                    COLUMNS[14]: clean_text(export_hs),
                    COLUMNS[15]: clean_text(orig_co),
                    COLUMNS[16]: issue_date,
                    COLUMNS[17]: clean_text(auth),
                    COLUMNS[18]: clean_text(qty),    
                    COLUMNS[19]: clean_text(uom),    
                    COLUMNS[20]: global_info["produced_in"],
                    COLUMNS[21]: global_info["exported_to"],
                    COLUMNS[22]: date_of_cert,
                    COLUMNS[23]: global_info["form_type"], 
                    COLUMNS[24]: global_usd,         
                    COLUMNS[25]: box13_val,          
                    COLUMNS[26]: global_third_party  
                })
    except Exception as e:
        return {"error": f"{file_name}: {str(e)}", "data": [], "file_name": file_name}
    return {"error": None, "data": extracted_data, "file_name": file_name}

# ==========================================
# 3. GIAO DIỆN STREAMLIT & ĐIỀU PHỐI UX/UI
# ==========================================
def select_local_folder():
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    folder_path = filedialog.askdirectory(master=root, title="Chọn thư mục chứa Form")
    root.destroy()
    return folder_path

def select_local_files():
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()
    file_paths = filedialog.askopenfilenames(master=root, title="Chọn các file PDF Form", filetypes=[("PDF Files", "*.pdf")])
    root.destroy()
    return file_paths

def init_session_state():
    if "pdf_files" not in st.session_state: st.session_state.pdf_files = []
    if "selection_mode" not in st.session_state: st.session_state.selection_mode = None 
    if "source_name" not in st.session_state: st.session_state.source_name = ""
    if "is_processing" not in st.session_state: st.session_state.is_processing = False
    if "extracted_data" not in st.session_state: st.session_state.extracted_data = None
    if "errors" not in st.session_state: st.session_state.errors = []

def reset_data_state():
    st.session_state.extracted_data = None
    st.session_state.errors = []

def main():
    st.set_page_config(page_title="Custom Form Extractor", layout="wide", page_icon="📑")
    init_session_state()
    
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
        button[kind="secondary"] {{ background-color: white; color: {DECATHLON_DARK}; border: 1px solid #D1D5DB; }}
        button[kind="secondary"]:hover {{ border-color: {DECATHLON_BLUE}; color: {DECATHLON_BLUE}; background-color: #F0F9FF; transform: translateY(-2px); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }}
        .stProgress > div > div > div > div {{ background-color: {DECATHLON_BLUE}; transition: width 0.3s ease; border-radius: 10px; height: 8px; }}
        .data-card {{ background: white; padding: 1.5rem; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); border: 1px solid #E5E7EB; }}
        .badge {{ display: inline-block; padding: 0.35rem 0.8rem; font-size: 0.85rem; font-weight: 600; border-radius: 9999px; background-color: #E0F2FE; color: #0369A1; margin-bottom: 1rem; }}
        .source-name {{ font-size: 0.9rem; color: #374151; font-weight: 600; background: #F3F4F6; padding: 6px 12px; border-radius: 6px; margin-top: 10px; margin-bottom: 5px; display: inline-block; word-break: break-all; }}
        div[data-testid="stExpander"] button {{ height: 32px !important; width: 32px !important; padding: 0 !important; font-size: 1.1rem !important; border: none !important; background: transparent !important; color: #9CA3AF !important; display: flex; align-items: center; justify-content: center; }}
        div[data-testid="stExpander"] button:hover {{ background-color: #FEE2E2 !important; color: #EF4444 !important; border-radius: 50% !important; transform: none !important; box-shadow: none !important; }}
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
            c_btn1, c_btn2 = st.columns(2)
            with c_btn1:
                if st.button("📁 Thư mục", disabled=st.session_state.is_processing):
                    folder = select_local_folder()
                    if folder:
                        st.session_state.pdf_files = glob.glob(os.path.join(folder, "*.pdf"))
                        st.session_state.selection_mode = "folder"
                        st.session_state.source_name = f"Thư mục: {os.path.basename(folder)}"
                        reset_data_state()
                        st.rerun()
            with c_btn2:
                if st.button("📄 Thêm Tệp lẻ", disabled=st.session_state.is_processing):
                    files = select_local_files()
                    if files:
                        if st.session_state.selection_mode == "folder":
                            st.session_state.pdf_files = list(files)
                        else:
                            current_files = set(st.session_state.pdf_files)
                            new_files = [f for f in files if f not in current_files]
                            st.session_state.pdf_files.extend(new_files)
                        st.session_state.selection_mode = "files"
                        st.session_state.source_name = "Danh sách tệp tùy chỉnh"
                        reset_data_state()
                        st.rerun()
            
            if st.session_state.source_name:
                st.markdown(f"<div class='source-name'>📂 {st.session_state.source_name}</div>", unsafe_allow_html=True)
                if len(st.session_state.pdf_files) > 0:
                    with st.expander("🛠️ Quản lý tệp (Bấm để xem/xóa)", expanded=True):
                        st.caption("Danh sách các file PDF chuẩn bị trích xuất:")
                        with st.container(height=200):
                            for f in st.session_state.pdf_files:
                                cf, cdel = st.columns([8.5, 1.5])
                                cf.markdown(f"📄 `{os.path.basename(f)}`")
                                if cdel.button("✖", key=f"del_{f}"):
                                    st.session_state.pdf_files.remove(f)
                                    if len(st.session_state.pdf_files) == 0:
                                        st.session_state.source_name = ""
                                        st.session_state.selection_mode = None
                                    st.rerun()
                    st.caption(f"✅ Sẵn sàng: **{len(st.session_state.pdf_files)}** file PDF")
                else:
                    st.markdown("<div style='color: #DC2626; font-size: 0.85rem; font-weight: 600; margin-top: 5px;'>⚠️ Không tìm thấy file PDF nào! Vui lòng chọn lại.</div>", unsafe_allow_html=True)
            elif not st.session_state.pdf_files:
                st.caption("ℹ️ Chưa có dữ liệu đầu vào.")

        with col2:
            st.markdown("**2. Quá trình xử lý**")
            has_files = len(st.session_state.pdf_files) > 0
            if has_files and st.session_state.extracted_data is None:
                if not st.session_state.is_processing:
                    if st.button("🚀 BẮT ĐẦU TRÍCH XUẤT", type="primary"):
                        st.session_state.is_processing = True
                        st.rerun()
                else:
                    if st.button("🛑 Hủy tiến trình", type="secondary"):
                        cancel_clicked = True
            else:
                st.button("🚀 BẮT ĐẦU TRÍCH XUẤT", type="primary", disabled=True)
            if st.session_state.is_processing:
                st.caption("⚡ Đang trích xuất dữ liệu...")
            elif not has_files and st.session_state.source_name:
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
                    file_name="Decathlon_Form_Extraction.xlsx",
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
            st.session_state.selection_mode = None
            st.session_state.is_processing = False
            st.session_state.extracted_data = None
            st.session_state.errors = []
            st.rerun()
            
        elif st.session_state.is_processing:
            st.markdown("<hr style='margin: 1.5rem 0; border-color: #F3F4F6;'>", unsafe_allow_html=True)
            progress_bar = st.progress(0)
            status_text = st.empty()
            all_extracted_data = []
            errors = []
            total_files = len(st.session_state.pdf_files)
            
            with ProcessPoolExecutor() as executor:
                futures = {executor.submit(process_single_pdf, path): path for path in st.session_state.pdf_files}
                for idx, future in enumerate(as_completed(futures)):
                    result = future.result()
                    if result["error"]: errors.append(result["error"])
                    else: all_extracted_data.extend(result["data"])
                    progress_bar.progress((idx + 1) / total_files)
                    status_text.info(f"⏳ Đã xử lý xong: `{result['file_name']}` ({idx + 1}/{total_files})")
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
