import gc
import streamlit as st
import pandas as pd
import os
import io
import re
import time
import pdfplumber
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- THƯ VIỆN BỔ SUNG CHO OCR ---
import pytesseract
from PIL import Image

# ==========================================
# 1. ĐỒNG BỘ CẤU HÌNH CỘT DỮ LIỆU ĐẦU RA MỚI KỲ VỌNG
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
    "Number, date of Invoices",
    "Invoice Number",            # Cột mới tách từ Box 10
    "Date of invoices",          # Cột mới tách từ Box 10
    "CARTON",
    "English description",
    "IMPORTING COUNTRY HS CODE",
    "EXPORTING COUNTRY HS CODE",
    "Original CO Reference Number",
    "Issuance Date",
    "Issuing Authority",
    "Quantity",                  # Tách riêng số lượng sản phẩm từ Box 7
    "UOM",                       # Tách riêng đơn vị tính từ Box 7
    "produced in",
    "exported to",
    "Date of certification",
    "Form", 
    "USD",                       # Chỉ lấy giá trị số USD từ Box 9
    "Third party",               # Cột dữ liệu lặp lại từ Box 7 của trang cuối
    "Box 13"
]

DECATHLON_BLUE = "#0082C3"
DECATHLON_DARK = "#1F2937"
BG_LIGHT = "#F9FAFB"

# ==========================================
# 2. HÀM CHUẨN HÓA ĐỊNH DẠNG NGÀY THÁNG (Định dạng: Date - Month - Year)
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

def extract_global_info(page):
    bbox_exporter = (30, 40, 290, 130)  
    bbox_consignee = (30, 130, 290, 200)
    bbox_transport = (30, 220, 290, 330)
    bbox_ref_no = (290, 40, 500, 110)
    box_movement = (38, 783, 55, 805)
    box_third_party = (162, 783, 180, 805)

    exporter = clean_text(page.crop(bbox_exporter).extract_text())
    consignee = clean_text(page.crop(bbox_consignee).extract_text())
    transport = clean_text(page.crop(bbox_transport).extract_text())
    
    raw_ref = page.crop(bbox_ref_no).extract_text()
    ref_match = re.search(r'Reference No\.\s*([A-Z0-9\-]+)', raw_ref, re.IGNORECASE)
    reference_no = ref_match.group(1) if ref_match else clean_text(raw_ref)

    def is_checked(bbox):
        cropped = page.crop(bbox)
        img = cropped.to_image(resolution=150).original.convert("L")
        try:
            pixels = list(img.get_flattened_data())
        except AttributeError:
            pixels = list(img.getdata())
        dark_pixels = sum(1 for p in pixels if p < 128)
        return "Yes" if (dark_pixels / len(pixels)) > 0.05 else "No"

    movement_cert = is_checked(box_movement)
    third_party = is_checked(box_third_party)

    bbox_box11 = (0, 545, 350, page.height)
    bbox_box12 = (300, 550, page.width, page.height)
    
    box11_text = clean_text(page.crop(bbox_box11).extract_text())
    box12_text = clean_text(page.crop(bbox_box12).extract_text())
    
    asean_china = ["CHINA", "VIETNAM", "MALAYSIA", "SINGAPORE", "INDONESIA", "THAILAND", "PHILIPPINES", "BRUNEI", "CAMBODIA", "LAOS", "MYANMAR"]
    country_matches = re.findall(r'\b(' + '|'.join(asean_china) + r')\b', box11_text, re.IGNORECASE)
    
    produced_in = country_matches[0].upper() if len(country_matches) > 0 else ""
    exported_to = country_matches[1].upper() if len(country_matches) > 1 else ""
    
    date_of_cert = ""
    date_match = re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', box12_text)
    if date_match:
        date_of_cert = format_to_dd_mm_yyyy(date_match.group(1))

    form_type = ""
    try:
        bbox_form = (290, 0, page.width, 150)
        cropped_form_page = page.crop(bbox_form)
        cropped_img_obj = cropped_form_page.to_image(resolution=150)
        pil_image = cropped_img_obj.original
        
        ocr_text = pytesseract.image_to_string(pil_image)
        form_match = re.search(r'FORM\s*([A-Za-z0-9]+)', ocr_text, re.IGNORECASE)
        if form_match:
            form_type = form_match.group(1).strip().upper()
    except Exception as e:
        print(f"[!] Cảnh báo OCR: Xảy ra lỗi khi đọc Form Type - {e}")
        form_type = "" 

    return {
        "exporter": exporter, "consignee": consignee, "transport": transport,
        "reference_no": reference_no, "movement_cert": movement_cert, "third_party": third_party,
        "produced_in": produced_in, "exported_to": exported_to, "date_of_cert": date_of_cert,
        "form_type": form_type
    }

def parse_description_fields(desc_text):
    text = re.sub(r'\s+', ' ', desc_text).strip()
    
    carton_match = re.search(r'([\d,\.]+)\s*CARTON', text, re.IGNORECASE)
    carton = carton_match.group(1).strip() if carton_match else ""
    
    qty, uom = "", ""
    qty_uom_match = re.search(r'-\s*([\d,\.]+)\s*([A-Za-z]+)', text)
    if qty_uom_match:
        qty = qty_uom_match.group(1).strip()
        uom = qty_uom_match.group(2).strip().upper()
        
    eng_desc = ""
    desc_match = re.search(r'CARTON\s*(.*?)\s*-\s*[\d,\.]+\s*[A-Za-z]+', text, re.IGNORECASE)
    if desc_match:
        eng_desc = desc_match.group(1).strip()
    else:
        desc_match = re.search(r'CARTON\s*(.*?)(?:IMPORTING|[\d,\.]+\s*(?:PCE|PR|SET|KGS))', text, re.IGNORECASE)
        if desc_match:
            eng_desc = desc_match.group(1).strip().strip("- ")
            
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
    
    return carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth

def extract_table_items(pdf):
    items = []
    current_item = None
    global_invoice = ""

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
    return items, global_invoice

def process_single_pdf(file_data):
    extracted_data = []
    
    file_name = file_data["name"]
    file_bytes = file_data["bytes"]
    
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            global_info = extract_global_info(pdf.pages[0])
            items, global_invoice = extract_table_items(pdf)
            
            box_13_list = []
            if global_info["third_party"] == "Yes": box_13_list.append("Third Party Invoicing")
            if global_info["movement_cert"] == "Yes": box_13_list.append("Movement Certificate")
            box_13_str = ", ".join(box_13_list)
            
            # ==========================================
            # SỬA LỖI CHÍNH: LÀM SẠCH TRIỆT ĐỂ DATA THIRD PARTY TRÊN MỌI CỘT BỊ ẢNH HƯỞNG
            # ==========================================
            third_party_column_val = ""
            for item in items:
                desc_text = item["desc"].strip()
                match = re.search(r'(?i)(Third\s+Party)', desc_text)
                if match:
                    # 1. Tách trích xuất giữ lại duy nhất giá trị Box 7 cho cột Third Party
                    raw_tp = desc_text[match.start():].strip()
                    third_party_column_val = re.split(r'(?i)TOTAL|USD|MYR|EUR', raw_tp)[0].strip()
                    
                    # 2. Xóa sạch dấu vết Third party dính ở Box 7 (Mô tả hàng hóa)
                    item["desc"] = desc_text[:match.start()].strip()
                    
                    # 3. FIX: Xóa sạch hoàn toàn text Third Party dính nhầm ở cột Origin Criteria (Box 8)
                    if item["origin"]:
                        item["origin"] = re.split(r'(?i)Third|DESIPRO|TOTAL', item["origin"])[0].strip()
                        
                    # 4. FIX: Xóa sạch hoàn toàn text Third Party dính nhầm ở cột Gross weight... (Box 9)
                    if item["weight_value"]:
                        item["weight_value"] = re.split(r'(?i)Third|DESIPRO', item["weight_value"])[0].strip()
                    break

            # Xử lý dồn các dòng CONTINUATION vào item thật liền trước
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
                
                # TÁCH DỮ LIỆU BOX 10 (Invoice Number & Date of invoices)
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
                
                carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth = parse_description_fields(desc)
                
                # CHUYỂN ĐỔI ĐỊNH DẠNG NGÀY THÁNG ĐỒNG BỘ (DD-MM-YYYY)
                issue_date = format_to_dd_mm_yyyy(issue_date)
                
                # BẢN VÁ LỖI TÊN BIẾN (qty & uom thay cho pce)
                if not qty:
                    qty_fallback = re.search(r'([\d,\.]+)\s*(PCE|PR|SET|KGS)\b', weight_value_text, re.IGNORECASE)
                    if qty_fallback:
                        qty = qty_fallback.group(1).strip()
                        uom = qty_fallback.group(2).strip().upper()
                
                # CHỈ LẤY ĐÚNG GIÁ TRỊ USD TẠI BOX 9
                usd_match = re.search(r'USD\s*([\d,\.]+)', weight_value_text, re.IGNORECASE)
                usd = usd_match.group(1).strip() if usd_match else ""
                
                if not usd:
                    usd_match_desc = re.search(r'USD\s*([\d,\.]+)', desc, re.IGNORECASE)
                    usd = usd_match_desc.group(1).strip() if usd_match_desc else ""
                
                # Làm sạch chuỗi Box 9 hiển thị (Chỉ lọc giữ lại thông tin USD, xóa MYR, EUR,...)
                weight_value_cleaned = weight_value_text
                weight_value_cleaned = re.sub(r'\b(MYR|EUR|SGD|VND)\s*[\d,\.]+', '', weight_value_cleaned, flags=re.IGNORECASE)
                weight_value_cleaned = re.sub(r'\s+', ' ', weight_value_cleaned).strip()

                item_no_val = clean_text(item["item_no"])
                if item_no_val.upper() == "CONTINUATION":
                    item_no_val = ""
                
                # Sắp xếp và map chuẩn xác dữ liệu đầu ra
                extracted_data.append({
                    COLUMNS[0]: global_info["exporter"],
                    COLUMNS[1]: global_info["consignee"],
                    COLUMNS[2]: global_info["transport"],
                    COLUMNS[3]: global_info["reference_no"],
                    COLUMNS[4]: item_no_val,
                    COLUMNS[5]: clean_text(item["marks"]),
                    COLUMNS[6]: desc,
                    COLUMNS[7]: origin_criteria_cleaned, # Đã làm sạch hoàn toàn text dính từ third party
                    COLUMNS[8]: weight_value_cleaned,     # Đã làm sạch hoàn toàn text dính từ third party
                    COLUMNS[9]: invoice,
                    COLUMNS[10]: invoice_number,      
                    COLUMNS[11]: invoice_date,        
                    COLUMNS[12]: clean_text(carton),
                    COLUMNS[13]: clean_text(eng_desc),
                    COLUMNS[14]: clean_text(import_hs),
                    COLUMNS[15]: clean_text(export_hs),
                    COLUMNS[16]: clean_text(orig_co),
                    COLUMNS[17]: clean_text(issue_date),
                    COLUMNS[18]: clean_text(auth),
                    COLUMNS[19]: clean_text(qty),
                    COLUMNS[20]: clean_text(uom),
                    COLUMNS[21]: global_info["produced_in"],
                    COLUMNS[22]: global_info["exported_to"],
                    COLUMNS[23]: global_info["date_of_cert"],
                    COLUMNS[24]: global_info["form_type"], 
                    COLUMNS[25]: usd,
                    COLUMNS[26]: third_party_column_val,  
                    COLUMNS[27]: box_13_str
                })
    except Exception as e:
        return {"error": f"{file_name}: {str(e)}", "data": [], "file_name": file_name}
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
            
            for idx, f_data in enumerate(safe_file_list):
                result = process_single_pdf(f_data)
                
                if result["error"]: 
                    errors.append(result["error"])
                else: 
                    all_extracted_data.extend(result["data"])
                
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
