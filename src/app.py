import gc
import streamlit as st
import pandas as pd
import os
import io
import re
import time
import requests  
import smtplib
import uuid
import shutil
import difflib
from collections import Counter
from html import escape as html_escape
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import numpy as np
from datetime import datetime, timedelta

# --- THƯ VIỆN BỔ SUNG CHO OCR ---
import pdfplumber
import fitz  # PyMuPDF
import pytesseract
from pytesseract import Output
from PIL import Image, ImageDraw, ImageOps
import logging

# 1. Configure logging level
logging.basicConfig(level=logging.INFO)

def _resolve_tesseract_cmd():
    """Tìm tesseract: biến môi trường TESSERACT_CMD -> PATH -> đường dẫn mặc định Linux/Windows.
    Trên server Linux (Streamlit Cloud) kết quả vẫn là '/usr/bin/tesseract' như trước."""
    env_cmd = os.environ.get("TESSERACT_CMD")
    if env_cmd:
        return env_cmd
    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in ("/usr/bin/tesseract", r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(candidate):
            return candidate
    return "/usr/bin/tesseract"

pytesseract.pytesseract.tesseract_cmd = _resolve_tesseract_cmd()
# ==========================================
# 1. CẤU HÌNH CỘT DỮ LIỆU ĐẦU RA 
# ==========================================
COLUMNS = [
    "Form", "Reference No", "Original CO Reference Number", "Item Number",
    "English description", "Quantity - Box 9", "UOM - Box 9", 
    "Quantity - Box 7", "UOM - Box 7",
    "USD", "Origin criteria (see Overleaf Notes)",
    "IMPORTING COUNTRY HS CODE", "EXPORTING COUNTRY HS CODE", "Invoice Number",
    "Date of invoices", "CARTON", "Original CO Issuance Date", "Issuing Authority",
    "Date of certification", "Products consigned from (Exporter's business name, address, country)",
    "Products consigned to (Consignee's name, address, country)", "Means of transport and route (as far as known)",
    "Produced in", "Exported to", "Marks and numbers on packages", "Box 13", 
    "Third party"
]

DECATHLON_BLUE = "#0082C3"
DECATHLON_DARK = "#1F2937"
BG_LIGHT = "#F9FAFB"

SHEET_NAME = "CO_Extract - Record"

def get_gspread_client():
    # Đọc credentials từ st.secrets khi deploy trên Streamlit Cloud
    # Hoặc nếu chạy local, Streamlit tự đọc từ .streamlit/secrets.toml
    if "gcp_service_account" in st.secrets:
        credentials = dict(st.secrets["gcp_service_account"])
        return gspread.service_account_from_dict(credentials)
    else:
        # Fallback cho trường hợp bạn test local mà vẫn để file json (nhớ gitignore nó)
        return gspread.service_account(filename="credential.json")

def record_access(session_id):
    try:
        gc_client = get_gspread_client()
        sh = gc_client.open(SHEET_NAME)
        worksheet = sh.worksheet("Access Record")
        
        # Lấy giờ server (UTC) cộng thêm 7 tiếng để ra giờ Việt Nam (GMT+7)
        vietnam_time = datetime.now() + timedelta(hours=7)
        login_time = vietnam_time.strftime("%Y-%m-%d %H:%M:%S")
        
        worksheet.append_row([session_id, login_time])
    except Exception as e:
        print(f"[!] Lỗi ghi nhận Access Record: {e}")

def record_file_processing(session_id, file_count):
    try:
        gc_client = get_gspread_client()
        sh = gc_client.open(SHEET_NAME)
        worksheet = sh.worksheet("File Record")
        
        # Lấy giờ server (UTC) cộng thêm 7 tiếng để ra giờ Việt Nam (GMT+7)
        vietnam_time = datetime.now() + timedelta(hours=7)
        action_time = vietnam_time.strftime("%Y-%m-%d %H:%M:%S")
        
        worksheet.append_row([session_id, action_time, file_count])
    except Exception as e:
        print(f"[!] Lỗi ghi nhận File Record: {e}")

MONTHS_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
    "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

def format_to_dd_mm_yyyy(date_str):
    """Hàm chuẩn hóa ngày DÙNG CHUNG cho cả Form D, E, AI.
    Mọi định dạng (01/11/2026, 2026-11-01, 10 November 2026, 30-SEP-2026...) -> chuỗi 'DD-MM-YYYY'.
    Không nhận diện được thì trả lại nguyên chuỗi (không đoán)."""
    if not date_str: return ""
    date_str = str(date_str).strip()
    match = re.search(r'(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})', date_str)
    if match: return f"{int(match.group(1)):02d}-{int(match.group(2)):02d}-{match.group(3)}"
    match = re.search(r'(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})', date_str)
    if match: return f"{int(match.group(3)):02d}-{int(match.group(2)):02d}-{match.group(1)}"
    match = re.search(r'(\d{1,2})\s*[\-\s]\s*([A-Za-z]+)\.?\s*[\-\s]\s*(\d{4})', date_str)
    if match:
        # So khớp theo 3 ký tự đầu của tên tháng: 'Sep', 'Sept', 'September' đều -> 09
        m_key = match.group(2).lower()[:3]
        if m_key in MONTHS_MAP: return f"{int(match.group(1)):02d}-{MONTHS_MAP[m_key]}-{match.group(3)}"
    return date_str

# Alias để code mới gọi rõ nghĩa hơn; vẫn là MỘT hàm duy nhất.
normalize_date = format_to_dd_mm_yyyy


# ==========================================
# 1b. CÁC HÀM DÙNG CHUNG: BOX 3 (TRANSPORT), BOX 11 (PRODUCED IN / EXPORTED TO)
# ==========================================
# Nhãn in sẵn trong Box 3 – chịu biến thể OCR: ’ ' ‘ ` ´, 'etc.' / 'ete.' / 'etc', thiếu/thừa khoảng trắng, ':' ';' '.'
TRANSPORT_LABEL_PATTERNS = [
    r"3\.?\s*Means\s+of\s+transport.*?\)",
    r"Departure\s*Date\s*[:;.]?",
    r"Vessel\s*[’'‘`´]?\s*s?\s*Na\w{1,3}\s*[/\\|lI1]?\s*(?:or\s*)?Air\s*[a-z]{3,6}\s*[,.]?\s*(?:et\s*[ce]\s*\.?)?\s*[:;.]?",
    r"(?<![A-Za-z])et[ce]\s*\.\s*[:;]",
    r"Port\s+of\s+Discharge\s*[:;.]?",
]

def strip_transport_labels(text):
    """Loại bỏ các nhãn in sẵn của Box 3, chỉ giữ nội dung thật (ngày, phương tiện, tàu, cảng)."""
    if not text: return ""
    for pat in TRANSPORT_LABEL_PATTERNS:
        text = re.sub(pat, " ", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip()


BOX11_PRODUCED_RE = re.compile(r"(?i)^produc")
BOX11_EXPORTED_RE = re.compile(r"(?i)^exported\b")
BOX11_COUNTRY_LABEL_RE = re.compile(r"(?i)^\(?c[o0]untr")
BOX11_IMPORTING_RE = re.compile(r"(?i)^\(?(?:im|lm|in)p?[o0]rt")
BOX11_LABEL_VOCAB_RE = re.compile(r"(?i)^\(?(?:country|importing|produced|exported|to|in|goods|the|for|and|place|date)\)?[.,:;]?$")
COUNTRY_VALUE_RE = re.compile(r"^[A-Z][A-Z .'\-]{2,40}$")

def _ocr_words_from_image(pil_img, scale, offset_x, offset_y, config=r'--oem 3 --psm 11'):
    """OCR một ảnh và trả về danh sách từ với tọa độ đã quy đổi về đơn vị point của PDF."""
    data = pytesseract.image_to_data(pil_img, output_type=Output.DICT, config=config)
    words = []
    for i, raw in enumerate(data['text']):
        txt = (raw or "").strip()
        if not txt: continue
        words.append({
            'text': txt,
            'x0': offset_x + data['left'][i] * scale,
            'x1': offset_x + (data['left'][i] + data['width'][i]) * scale,
            'top': offset_y + data['top'][i] * scale,
            'bottom': offset_y + (data['top'][i] + data['height'][i]) * scale,
        })
    return words

def _box11_region(page_width, page_height):
    # Box 11 là nửa trái, nằm giữa đáy bảng hàng hóa và Box 13 (tỷ lệ theo khổ A4 của mẫu C/O)
    return (0, page_height * 0.62, page_width * 0.5, page_height * 0.92)

def _first_label(words, regex, min_top=None):
    cands = [w for w in words if regex.search(w['text']) and (min_top is None or w['top'] > min_top)]
    return min(cands, key=lambda w: w['top']) if cands else None

def _value_line_below(label, stop_label, value_words):
    """Lấy dòng chữ nằm NGAY DƯỚI nhãn (và phía trên nhãn kế tiếp nếu có)."""
    lh = max(label['bottom'] - label['top'], 4)
    min_top = label['top'] + 0.6 * lh
    max_top = (stop_label['top'] + 0.5 * max(stop_label['bottom'] - stop_label['top'], 4)) if stop_label else label['bottom'] + 30
    cands = [w for w in value_words
             if min_top <= w['top'] < max_top and not BOX11_LABEL_VOCAB_RE.match(w['text'])]
    if not cands: return ""
    first_top = min(w['top'] for w in cands)
    line = sorted([w for w in cands if abs(w['top'] - first_top) <= 3], key=lambda w: w['x0'])
    # Tách cụm theo khoảng trống ngang lớn: chữ trên con dấu/chữ ký cùng hàng không bị ghép vào tên nước
    clusters, current = [], [line[0]]
    for prev, w in zip(line, line[1:]):
        if w['x0'] - prev['x1'] > 12:
            clusters.append(current); current = [w]
        else:
            current.append(w)
    clusters.append(current)
    for cluster in clusters:
        value = " ".join(w['text'] for w in cluster)
        value = re.sub(r"[^A-Za-z .'\-]", "", value).strip(" .-'").upper()
        value = re.sub(r"\s+", " ", value)
        if COUNTRY_VALUE_RE.match(value):
            return value
    return ""

def locate_box11_countries(label_words, value_words):
    """
    Xác định 'Produced in' và 'Exported to' theo VỊ TRÍ so với nhãn của Box 11:
      - Produced in  = dòng ngay dưới nhãn '...produced in' (trên nhãn '(Country)')
      - Exported to  = dòng ngay dưới nhãn '...goods exported to' (trên nhãn '(Importing Country)')
    Trả về (produced_in, exported_to, warnings). Không xác định được thì để trống, không đoán.
    """
    warnings = []
    produced_in, exported_to = "", ""
    lbl_prod = _first_label(label_words, BOX11_PRODUCED_RE)
    if lbl_prod:
        lbl_country = _first_label(label_words, BOX11_COUNTRY_LABEL_RE, min_top=lbl_prod['bottom'])
        produced_in = _value_line_below(lbl_prod, lbl_country, value_words)
    lbl_exp = _first_label(label_words, BOX11_EXPORTED_RE, min_top=lbl_prod['bottom'] if lbl_prod else None)
    if lbl_exp:
        lbl_imp = _first_label(label_words, BOX11_IMPORTING_RE, min_top=lbl_exp['bottom'])
        exported_to = _value_line_below(lbl_exp, lbl_imp, value_words)

    if not lbl_prod: warnings.append("Box 11: không tìm thấy nhãn 'produced in' → để trống Produced in")
    elif not produced_in: warnings.append("Box 11: không đọc được tên nước dưới nhãn 'produced in' → để trống Produced in")
    if not lbl_exp: warnings.append("Box 11: không tìm thấy nhãn 'goods exported to' → để trống Exported to")
    elif not exported_to: warnings.append("Box 11: không đọc được tên nước dưới nhãn 'exported to' → để trống Exported to")
    return produced_in, exported_to, warnings

def extract_box11_countries_textlayer(page):
    """PDF có text layer (pdfplumber page). Giá trị lấy từ text layer; nhãn lấy từ text layer,
    nếu nhãn chỉ nằm trong ảnh nền của mẫu in thì OCR vùng Box 11 để định vị nhãn."""
    bbox = _box11_region(page.width, page.height)
    region = page.crop(bbox)
    text_words = region.extract_words()
    label_words = text_words
    if not (_first_label(text_words, BOX11_PRODUCED_RE) and _first_label(text_words, BOX11_EXPORTED_RE)):
        try:
            img = region.to_image(resolution=200).original
            label_words = _ocr_words_from_image(img, 72 / 200, bbox[0], bbox[1])
        except Exception as e:
            return "", "", [f"Box 11: lỗi OCR khi tìm nhãn ({e}) → để trống Produced in / Exported to"]
    produced_in, exported_to, warnings = locate_box11_countries(label_words, text_words)
    # Nếu tên nước không có trong text layer (vd. đóng dấu/viết tay) thì đọc giá trị từ OCR
    if (not produced_in or not exported_to) and label_words is not text_words:
        p2, e2, _ = locate_box11_countries(label_words, label_words)
        if not produced_in and p2: produced_in = _snap_ocr_country(p2, "Produced in", warnings)
        if not exported_to and e2: exported_to = _snap_ocr_country(e2, "Exported to", warnings)
        warnings = [w for w in warnings
                    if not (("Produced in" in w and produced_in) or ("Exported to" in w and exported_to))]
    return produced_in, exported_to, warnings

KNOWN_COUNTRIES = [
    "BRUNEI DARUSSALAM", "BRUNEI", "CAMBODIA", "INDONESIA", "LAO PDR", "LAOS", "MALAYSIA", "MYANMAR",
    "PHILIPPINES", "SINGAPORE", "THAILAND", "VIETNAM", "VIET NAM", "CHINA", "INDIA", "JAPAN", "KOREA",
    "AUSTRALIA", "NEW ZEALAND", "HONG KONG",
]

def _snap_ocr_country(value, field, warnings):
    """Chỉ dùng cho giá trị đọc bằng OCR: sửa lỗi OCR nhỏ (vd. 'VIETNA' -> 'VIETNAM') theo danh sách nước.
    Không khớp đủ gần thì giữ nguyên giá trị OCR và cảnh báo để người dùng kiểm tra."""
    if not value: return value
    if value in KNOWN_COUNTRIES: return value
    match = difflib.get_close_matches(value, KNOWN_COUNTRIES, n=1, cutoff=0.8)
    if match: return match[0]
    warnings.append(f"Box 11: {field} = '{value}' (đọc bằng OCR, không có trong danh sách nước) – vui lòng kiểm tra")
    return value

def extract_box11_countries_scanned(fitz_page, zoom=300 / 72):
    """PDF scan (PyMuPDF page): OCR vùng Box 11, nhãn và giá trị đều lấy từ OCR."""
    x0, y0, x1, y1 = _box11_region(fitz_page.rect.width, fitz_page.rect.height)
    pix = fitz_page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=fitz.Rect(x0, y0, x1, y1))
    mode = "RGBA" if pix.alpha else "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples).convert("RGB")
    words = _ocr_words_from_image(img, 1 / zoom, x0, y0)
    produced_in, exported_to, warnings = locate_box11_countries(words, words)
    produced_in = _snap_ocr_country(produced_in, "Produced in", warnings)
    exported_to = _snap_ocr_country(exported_to, "Exported to", warnings)
    return produced_in, exported_to, warnings


# ==========================================
# 1c. BOX 7 (QUANTITY/UOM TRONG MÔ TẢ HÀNG) CHO FILE SCAN
# ==========================================
# Nếu Box 7 của C/O hoàn toàn KHÔNG ghi số lượng (vd. chỉ có "2 CARTON / BOY'S T-SHIRT"),
# lấy số lượng/đơn vị của Box 9 cho cột H/I và gắn cảnh báo. Đặt False để để trống.
BOX7_FALLBACK_TO_BOX9 = True
UOM_WORDS_RE = r'(?:PCE|PCS|PC|PIECES?|PRS?|PAIRS?|SETS?|UNITS?|NOS|KGM|KGS?|DOZ|DZN|BOX(?:ES)?|ROLLS?|PKGS?|MTR|LTR)'

def extract_box7_qty_scanned(raw_desc, block_text):
    """Các cách trình bày Box 7 đã gặp:
       (1) 'SOCKS - 100 SET'           (2) '... 100 SET' ngay trước 'IMPORTING COUNTRY'
       (3) '... 100 SET' ở cuối mô tả  (4) gạch nối bị OCR đọc thành ~ = – — hoặc mất khoảng trắng
       (5) '100 SET' nằm riêng một dòng (giống Form D) – có thể dính 'USD xx' của Box 9 cùng hàng."""
    m = re.search(r'[-–—]\s*(\d+[\d\.,]*)\s*([A-Za-z]{2,8})\b', raw_desc)
    if m: return m.group(1).replace(',', ''), m.group(2).upper()
    m = re.search(r'(\d+[\d\.,]*)\s*([A-Za-z]{2,8})\s*(?:\n|\s)*IMPORTING\s+COUNTRY', block_text, re.IGNORECASE)
    if m: return m.group(1).replace(',', ''), m.group(2).upper()
    m = re.search(r'(?:\s|^)(\d+[\d\.,]*)\s*([A-Za-z]{2,8})\s*$', raw_desc)
    if m: return m.group(1).replace(',', ''), m.group(2).upper()
    # (4) gạch nối bị OCR đọc sai
    m = re.search(r'[~=]\s*(\d+[\d\.,]*)\s*(' + UOM_WORDS_RE + r')\b', raw_desc, re.IGNORECASE)
    if m: return m.group(1).replace(',', ''), m.group(2).upper()
    # (5) số lượng nằm riêng một dòng trong vùng mô tả (trước IMPORTING COUNTRY)
    desc_zone = re.split(r'(?i)(?:IMPORTING\s+COUNTRY|Original\s+CO|TOTAL|Page)', block_text)[0]
    for line in desc_zone.split('\n'):
        if re.search(r'N/M|\bCARTONS?\b', line, re.I): continue
        rest = re.sub(r'(?i)USD\s*[\d\.,]+', ' ', line)
        rest = re.sub(r'\b\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4}\b', ' ', rest)
        rest = re.sub(r'(?i)\b[VY]N\s*\d{6,}\S*', ' ', rest)
        rest = re.sub(r'[|_\[\]]', ' ', rest).strip()
        m = re.fullmatch(r'(\d+[\d\.,]*)\s*(' + UOM_WORDS_RE + r')\.?', rest, re.IGNORECASE)
        if m: return m.group(1).replace(',', ''), m.group(2).upper()
    return "", ""

def clean_text(text):
    if not text: return ""
    text = re.sub(r'(?i)\b(?:page\s*\d+\s*of\s*\d+|page\s*\d+\s*of|\d+\s*of\s*\d+|\d+\s*of|page\s*\d+|of\s*\d+)\b', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'^:\s*', '', text)
    return text.strip() if re.search(r'[A-Za-z0-9]', text) else ""

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
def remove_vertical_table_lines(pil_img):
    """
    [Tầng 2 - Computer Vision]: Xóa đường kẻ dọc của bảng C/O đè lên chữ 'IN' hoặc ',VN'
    mà không làm đứt nét ngang/chéo của chữ cái.
    """
    gray = np.array(pil_img.convert("L"))
    dark_mask = (gray < 200).astype(np.int32)
    
    # Cửa sổ trượt dọc 130px (~1cm ở 300 DPI): Đường kẻ dọc liên tục > 80%, chữ cái luôn có khoảng trắng dòng
    kernel_h = 65
    padded_v = np.pad(dark_mask, ((kernel_h, kernel_h), (0, 0)), mode='constant')
    cumsum_v = np.cumsum(padded_v, axis=0)
    moving_sum_v = cumsum_v[2 * kernel_h:, :] - cumsum_v[:-2 * kernel_h, :]
    is_vline = moving_sum_v >= int(2 * kernel_h * 0.8)
    
    # Kiểm tra độ dày ngang (chỉ xóa vạch kẻ mảnh <= 5px, giữ lại nét chữ cái giao nhau với vạch kẻ)
    padded_h = np.pad(dark_mask, ((0, 0), (4, 4)), mode='constant')
    cumsum_h = np.cumsum(padded_h, axis=1)
    h_sum = cumsum_h[:, 8:] - cumsum_h[:, :-8]
    
    thin_vline = is_vline & (h_sum <= 5)
    erase_mask = thin_vline.copy()
    erase_mask[:, 1:] |= thin_vline[:, :-1]
    erase_mask[:, :-1] |= thin_vline[:, 1:]
    
    cleaned = gray.copy()
    cleaned[erase_mask] = 255
    return Image.fromarray(cleaned).convert("RGB")


def parse_col10_invoice_block(col10_text, form_type=""):
    """
    [Tầng 3 - NLP Parser trên vùng Cột 10 đã cô lập]:
    Xử lý chính xác 1 hoặc nhiều Invoice (kể cả bị ngắt dòng như VN26000174,VN / 26000175)
    và chỉ thêm ' IN' khi ngay sau số hóa đơn trên cùng dòng có hậu tố chữ (IN, LN, 1N, |N...).
    """
    if not col10_text:
        return "", ""
        
    # 1. Trích xuất Date of Invoice (DD/MM/YYYY)
    date_inv = ""
    d_match = re.search(r'\b(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})\b', col10_text)
    if d_match:
        date_inv = d_match.group(1)
        col10_text = col10_text.replace(date_inv, '')

    cleaned_invs = []
    has_letter_suffix = False

    # 2. Duyệt từng dòng trong Cột 10
    for line in col10_text.split('\n'):
        line_clean = line.strip()
        if not line_clean:
            continue
        # Bỏ qua dòng tiêu đề cột hoặc chữ từ Box 12 nếu vô tình chạm biên
        if re.search(r'(?i)\b(Number|date|invoices?|certified|basis|control|carried|out)\b', line_clean):
            continue

        # [MỚI] Nếu chữ IN/LN bị rớt xuống thành 1 dòng riêng biệt (như Item 8 của file 010237)
        if re.match(r'^[\s\|\!\[\]\/]*(?:[I1L\|\!\[\]\/JT]N|IN|LN)[\s\.\,\|\!\]]*$', line_clean, re.IGNORECASE):
            has_letter_suffix = True
            continue

        # Bắt mã số hóa đơn: Dùng (?<![0-9]) thay vì \b để bắt được cả "VN26000174" lẫn dòng rớt "26000175"
        pattern = r'(?<![0-9])(?:[VY]N[ \t]*)?(2[4-9]\d{6}(?:\-\d+)?)(?![0-9])([ \t]*[^\n\r,;]*)?'
        for m in re.finditer(pattern, line_clean, re.IGNORECASE):
            num_part = m.group(1)
            rest_segment = (m.group(2) or "").strip()
            inv_code = "VN" + re.sub(r'\s+', '', num_part)
            if inv_code not in cleaned_invs:
                cleaned_invs.append(inv_code)

            # Kiểm tra hậu tố chữ ngay phía sau số hóa đơn trên cùng dòng (VD: "IN", "LN", "1N", "|N", "[N")
            if rest_segment:
                s = re.sub(r'^[\s\.\,\;\:\-\_]+', '', rest_segment)
                # Nếu bắt đầu bằng V hoặc Y (như ",VN") thì đó là tiền tố của hóa đơn thứ 2, KHÔNG phải hậu tố IN
                if s and not re.match(r'^[VYvy]', s):
                    tok_m = re.match(r'^([A-Za-z0-9\|!\[\]\/]{1,3})(?=\s|$|[\,\.\;\:\|\!\]])', s)
                    if tok_m:
                        tok = tok_m.group(1).upper()
                        if re.match(r'^(?:[I1L\|\!\[\]\/JT][NMHUI1L\|]{1,2}|[A-Z1\|!\[\]\/]{1,2}N)$', tok):
                            has_letter_suffix = True

    # Fallback nếu mã hóa đơn có định dạng số khác
    if not cleaned_invs:
        fallback_matches = re.findall(r'[VY]N\s*([0-9]+(?:\-[0-9]+)?)', col10_text, re.IGNORECASE)
        for fb in fallback_matches:
            inv_code = "VN" + re.sub(r'\s+', '', fb)
            if inv_code not in cleaned_invs:
                cleaned_invs.append(inv_code)

    if cleaned_invs:
        invoice_str = ",".join(cleaned_invs)
        if form_type == "AI" and has_letter_suffix:
            invoice_str += " IN"
    else:
        invoice_str = ""

    return invoice_str, date_inv


def extract_page_items_and_col10(page, mat, tess_config, form_type):
    """
    Trích xuất đồng thời văn bản Cột 5-9 và ghép cặp tọa độ Y với Cột 10 trên từng trang PDF Scan.
    Giới hạn đáy Cột 10 ở 575 pt để cắt bỏ hoàn toàn chữ 'ed out,' của Box 12.
    """
    y1 = 300
    y2 = page.rect.height - 180
    if y2 <= y1:
        return "", []

    # 1. Cắt vùng Cột 5 -> Cột 9 (x từ 0 đến 488 pt)
    left_rect = fitz.Rect(0, y1, 488, y2)
    pix_left = page.get_pixmap(matrix=mat, clip=left_rect)
    img_left = Image.frombytes("RGB", [pix_left.width, pix_left.height], pix_left.samples) if not pix_left.alpha else Image.frombytes("RGBA", [pix_left.width, pix_left.height], pix_left.samples).convert("RGB")
    img_left_clean = remove_vertical_table_lines(img_left)
    
    page_main_text = pytesseract.image_to_string(img_left_clean, config=tess_config)
    left_data = pytesseract.image_to_data(img_left_clean, output_type=Output.DICT, config=tess_config)
    
    nm_y_positions = []
    for idx, word in enumerate(left_data['text']):
        w_clean = word.strip()
        if re.search(r'(?i)^(?:N/M|N/W|M/N|N\.M)$', w_clean):
            nm_y_positions.append(left_data['top'][idx])

    # 2. Cắt riêng vùng Cột 10: Giới hạn đáy ở 575 pt (trên đường kẻ đáy bảng 580 pt) để KHÔNG dính Box 12 ("ed out,")
    y2_col10 = min(575, y2)
    col10_rect = fitz.Rect(485, y1, min(570, page.rect.width), y2_col10)
    pix_col10 = page.get_pixmap(matrix=mat, clip=col10_rect)
    img_col10 = Image.frombytes("RGB", [pix_col10.width, pix_col10.height], pix_col10.samples) if not pix_col10.alpha else Image.frombytes("RGBA", [pix_col10.width, pix_col10.height], pix_col10.samples).convert("RGB")
    img_col10_clean = remove_vertical_table_lines(img_col10)
    
    col10_data = pytesseract.image_to_data(img_col10_clean, output_type=Output.DICT, config=r'--oem 3 --psm 6')
    
    # Gom các từ trong Cột 10 theo tọa độ dòng Y
    col10_lines = []
    line_groups = {}
    for idx, word in enumerate(col10_data['text']):
        w_clean = word.strip()
        if not w_clean:
            continue
        y_top = col10_data['top'][idx]
        x_left = col10_data['left'][idx]
        matched_y = None
        for ky in line_groups:
            if abs(y_top - ky) <= 15:
                matched_y = ky
                break
        if matched_y is None:
            line_groups[y_top] = [(x_left, w_clean)]
        else:
            line_groups[matched_y].append((x_left, w_clean))

    for ky in sorted(line_groups.keys()):
        sorted_words = [w for _, w in sorted(line_groups[ky], key=lambda item: item[0])]
        col10_lines.append((ky, " ".join(sorted_words)))

    # 3. Phân cụm các dòng Cột 10 về từng Item dựa trên tọa độ Y của "N/M"
    nm_regex_matches = list(re.finditer(r'(?i)(?:N/M|N\s*/\s*M|N/W|M/N|N\.M)', page_main_text))
    num_items_on_page = len(nm_regex_matches)
    
    page_col10_results = []
    if num_items_on_page > 0:
        # Lọc bỏ phần tiêu đề "10. Number and date of invoices" (nằm ở y_pos <= 360 px) và phần dưới đáy bảng (> 1140 px)
        valid_col10_lines = [(y_pos, txt) for y_pos, txt in col10_lines if 360 < y_pos <= 1140]
        
        if num_items_on_page == 1:
            inv_str, dt_str = parse_col10_invoice_block("\n".join([txt for _, txt in valid_col10_lines]), form_type)
            page_col10_results.append((inv_str, dt_str))
        else:
            if len(nm_y_positions) == num_items_on_page:
                anchors = sorted(nm_y_positions)
            else:
                vn_ys = [y_pos for y_pos, txt in valid_col10_lines if re.search(r'[VY]N|2[4-9]\d{6}', txt, re.I)]
                anchors = vn_ys[:num_items_on_page] if len(vn_ys) >= num_items_on_page else [420 + k * 350 for k in range(num_items_on_page)]

            item_buckets = [[] for _ in range(num_items_on_page)]
            for y_pos, txt in valid_col10_lines:
                assigned_idx = 0
                for idx_a in range(num_items_on_page):
                    if y_pos >= anchors[idx_a] - 60:
                        assigned_idx = idx_a
                item_buckets[assigned_idx].append(txt)

            for bucket in item_buckets:
                inv_str, dt_str = parse_col10_invoice_block("\n".join(bucket), form_type)
                page_col10_results.append((inv_str, dt_str))

    return page_main_text, page_col10_results



def process_scanned_pdf(pdf, file_bytes, file_name):
    extracted_data = []
    
    exporter = "DECATHLON LOGISTICS MALAYSIA SDN. BHD.\nPLOT D40 & D44\nJALAN DPB/8, ZONE B\nPELABUHAN TANJUNG PELEPAS\n81560 GELANG PATAH, JOHOR, MALAYSIA"
    consignee = "DECATHLON VIETNAM CO., LTD\nPAX SKY BUILDING, 5TH FLOOR\n26 UNG VAN KHIEM, WARD 25,\nBINH THANH DISTRICT\n700000 HO CHI MINH CITY VIETNAM"
    reference_no = file_name.replace(".pdf", "")
    
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        page_first = doc[0]
        zoom = 300 / 72
        mat = fitz.Matrix(zoom, zoom)
        tess_config = r'--oem 3 --psm 6'

        # ---------------------------------------------------------
        # GLOBAL INFO (ƯU TIÊN LẤY FORM TỪ TÊN FILE REFERENCE NO)
        # ---------------------------------------------------------
        top_rect = fitz.Rect(0, 0, page_first.rect.width, 300)
        pix_top = page_first.get_pixmap(matrix=mat, clip=top_rect)
        img_top = Image.frombytes("RGB", [pix_top.width, pix_top.height], pix_top.samples) if not pix_top.alpha else Image.frombytes("RGBA", [pix_top.width, pix_top.height], pix_top.samples).convert("RGB")
        top_text = pytesseract.image_to_string(img_top, config=tess_config).strip()
        top_flat = re.sub(r'[\n\|]', ' ', top_text)

        form_type = ""
        # 1. Ưu tiên trích xuất trực tiếp từ tên file (VD: KL-2026-AI-21-010237 -> AI, KL-2026-E-13-70037 -> E, KL-202609-CCF-... -> D)
        fn_match = re.search(r'(?i)\bKL-\d{4,6}-([A-Za-z0-9]+)-', reference_no)
        if fn_match:
            raw_ft = fn_match.group(1).upper()
            if raw_ft in ["AI", "AL", "A1"]:
                form_type = "AI"
            elif raw_ft in ["CCF", "D"]:
                form_type = "D"
            else:
                form_type = raw_ft
        else:
            fn_fallback = re.search(r'(?i)[\-_](AI|A1|AL|CCF|E|D|AK|AJ|AANZ)[\-_]', reference_no)
            if fn_fallback:
                raw_ft = fn_fallback.group(1).upper()
                if raw_ft in ["AI", "AL", "A1"]:
                    form_type = "AI"
                elif raw_ft == "CCF":
                    form_type = "D"
                else:
                    form_type = raw_ft

        # 2. Fallback bằng OCR trên trang 1 nếu tên file bị đặt khác chuẩn
        if not form_type:
            form_match = re.search(r'(?i)F[O0]R[MN]\s*([A-Za-z0-9\|]+)', top_flat)
            if form_match:
                form_type = form_match.group(1).upper()
                if form_type in ["AL", "A1", "A|", "A L", "A I"]:
                    form_type = "AI"
            elif re.search(r'(?i)ASEAN\s*[\-\s]\s*INDIA', top_flat):
                form_type = "AI"

        produced_in, exported_to = "", ""
        file_warnings = []
        if form_type == "AI":
            produced_in = "INDIA"
            exported_to = "VIETNAM"
        elif form_type == "D":
            # Form D: xác định theo vị trí so với nhãn Box 11 trên trang cuối (không đọc vùng đầu trang)
            try:
                produced_in, exported_to, box11_warn = extract_box11_countries_scanned(doc[-1])
                file_warnings.extend(box11_warn)
            except Exception as e:
                file_warnings.append(f"Box 11: lỗi OCR ({e}) → để trống Produced in / Exported to")
        else:
            prod_match = re.search(r'([A-Za-z\s]+)\s*\(Country\)', top_flat, re.IGNORECASE)
            if prod_match: produced_in = prod_match.group(1).strip().upper()
            exp_to_match = re.search(r'exported to(.*?)\(Importing Country\)', top_flat, re.IGNORECASE)
            if exp_to_match:
                raw_exp = exp_to_match.group(1)
                caps = re.findall(r'\b[A-Z]{3,}\b', raw_exp)
                exported_to = caps[-1] if caps else raw_exp.strip().upper()

        # ---------------------------------------------------------
        # BOX 3
        # ---------------------------------------------------------
        box3_rect = fitz.Rect(0, 160, 300, 315)
        pix3 = page_first.get_pixmap(matrix=mat, clip=box3_rect)
        img3 = Image.frombytes("RGB", [pix3.width, pix3.height], pix3.samples) if not pix3.alpha else Image.frombytes("RGBA", [pix3.width, pix3.height], pix3.samples).convert("RGB")
        
        b3_clean = pytesseract.image_to_string(img3, config=tess_config).strip()
        b3_clean = re.sub(r'\s+', ' ', b3_clean).strip()
        b3_clean = strip_transport_labels(b3_clean)

        transport = ""
        transport_match = re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4}.*)', b3_clean, re.IGNORECASE)
        if transport_match: transport = transport_match.group(1).strip()
        else: transport = b3_clean.strip()
        transport = strip_transport_labels(transport)

        # ---------------------------------------------------------
        # DATE OF CERTIFICATION
        # ---------------------------------------------------------
        cert_rect = fitz.Rect(0, 550, page_first.rect.width, page_first.rect.height)
        pix_cert = page_first.get_pixmap(matrix=mat, clip=cert_rect)
        img_cert = Image.frombytes("RGB", [pix_cert.width, pix_cert.height], pix_cert.samples) if not pix_cert.alpha else Image.frombytes("RGBA", [pix_cert.width, pix_cert.height], pix_cert.samples).convert("RGB")
        cert_text = pytesseract.image_to_string(img_cert, config=tess_config).strip()
        
        date_cert = ""
        cert_match = re.search(r'Lumpur\s*,\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})', cert_text, re.IGNORECASE)
        if cert_match:
            date_cert = normalize_date(cert_match.group(1).strip())

        # ---------------------------------------------------------
        # MAIN ITEMS + SPATIAL COLUMN 10 EXTRACTION
        # ---------------------------------------------------------
        main_text = ""
        all_col10_items = []
        for page in doc:
            page_text, page_col10 = extract_page_items_and_col10(page, mat, tess_config, form_type)
            if page_text:
                main_text += "\n" + page_text
                all_col10_items.extend(page_col10)

    logging.info(main_text)
    box_13_str = extract_box13_scanned(pdf.pages[-1])
    
    tp_match = re.search(r'(?i)(?:Third\s+Party\s+Invoicing\s+by|Third\s+Party|Third\s+Country)\s*[:;]?\s*(.*?)(?=\bTOTAL\b|\bPage\b|$)', main_text, re.DOTALL)
    third_party_val = clean_text(tp_match.group(1)) if tp_match else ""
    
    if form_type == "AI":
        third_party_val = "DESIPRO PTE LTD 230 STADIUM BOULEVARD 397799 SINGAPORE"

    matches = list(re.finditer(r'(?i)(?:N/M|N\s*/\s*M|N/W|M/N|N\.M)', main_text))
    
    if not matches:
        # Ghi theo TÊN cột (trước đây dùng chỉ số cũ nên các giá trị bị lệch 2 cột)
        error_row = {col: "" for col in COLUMNS}
        error_row.update({
            "Form": form_type, "Reference No": reference_no,
            "English description": "[Lỗi OCR] Không tìm thấy dữ liệu Item.",
            "Date of certification": date_cert,
            "Products consigned from (Exporter's business name, address, country)": exporter,
            "Products consigned to (Consignee's name, address, country)": consignee,
            "Means of transport and route (as far as known)": transport,
            "Produced in": produced_in, "Exported to": exported_to,
            "Box 13": box_13_str, "Third party": third_party_val,
        })
        extracted_data.append(error_row)
        file_warnings.append("Không tìm thấy dòng hàng (N/M) bằng OCR – kiểm tra chất lượng bản scan")
    else:
        for i in range(len(matches)):
            start = matches[i].start()
            end = matches[i+1].start() if i + 1 < len(matches) else len(main_text)
            block = main_text[start:end]
            item_no = str(i + 1)

            block_fixed = re.sub(r"(?<![A-Za-z'’])\bS(?=\s+[A-Z]{2,5}\b)", "5", block)
            block_fixed = re.sub(r"(?<=\d)S(?=\s+[A-Z]{2,5}\b)", "5", block_fixed)
            block_fixed = re.sub(r"(?<=\s)S(?=\s+[\d\.\,\+])", "5", block_fixed)

            qty, uom, usd, date_inv = "", "", "", ""
            carton, invoice, origin, desc, orig_co, orig_date = "", "", "", "", "", ""

            # --- CARTON ---
            c_match = re.search(r'(\d+)[A-Za-z]*\s*CARTON', block_fixed, re.IGNORECASE)
            carton = c_match.group(1) if c_match else ""

            # --- INVOICE NUMBER & DATE OF INVOICE (LẤY TỪ SPATIAL CỘT 10) ---
            if i < len(all_col10_items):
                invoice, date_inv = all_col10_items[i]

            # Fallback dự phòng nếu ngày hoặc mã hóa đơn chưa bắt được từ Cột 10
            if not date_inv:
                d_match = re.search(r'\b(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})\b', block_fixed)
                if d_match:
                    date_inv = d_match.group(1)

            if not invoice:
                invoice, _ = parse_col10_invoice_block(block_fixed, form_type)

            # --- QUANTITY & UOM ---
            qu_match = re.search(r'\b(\d+[\d\.,]*)\s+([A-Z]{2,5})\b', block_fixed)
            if qu_match:
                qty = qu_match.group(1).replace(',', '')
                uom = qu_match.group(2)

            # --- USD ---
            usd_match = re.search(r'USD\s*([\d\.,]+)', block_fixed, re.IGNORECASE)
            if usd_match:
                usd = usd_match.group(1).replace(',', '')
            else:
                if date_inv:
                    fb_match = re.search(r'\b(\d+[\.,]\d{2})\s*(?:[A-Za-z\s]{0,10})?' + re.escape(date_inv), block_fixed)
                    if fb_match: usd = fb_match.group(1).replace(',', '.')
                if not usd:
                    decimals = re.findall(r'\b(\d+\.\d{2})\b', block_fixed)
                    for dec in decimals:
                        if dec != qty:
                            usd = dec; break

            # --- HS CODE ---
            imp_m = re.search(r'(?i)IMPORTING\s+COUNTRY\s+HS\s+CODE\s*[:\-]?\s*(\d{8,10})', block_fixed)
            imp_hs = imp_m.group(1)[:8] if imp_m else ""

            exp_m = re.search(r'(?i)EXPORTING\s+COUNTRY\s+HS\s+CODE\s*[:\-]?\s*(\d{8,10})', block_fixed)
            exp_hs = exp_m.group(1)[:8] if exp_m else ""

            # --- ORIGINAL CO REFERENCE NUMBER ---
            orig_match = re.search(r'(?is)CO\s+Reference\s+Number[\s:,\-\|_]*\n*(.*?)(?=\n*Issuance|\n*Date|\n*Page|\n*TOTAL|$)', block_fixed)
            if orig_match:
                orig_co_raw = orig_match.group(1).strip()
                orig_co = re.sub(r'\s+', '', orig_co_raw)
                # Cắt bỏ mọi ký tự rác (như || hoặc ||_) đứng trước chữ TEXCOM hoặc ở 2 đầu chuỗi
                if "TEXCOM" in orig_co.upper():
                    orig_co = orig_co[orig_co.upper().find("TEXCOM"):]
                orig_co = re.sub(r'^[^A-Za-z0-9]+|[^A-Za-z0-9]+$', '', orig_co)
                
                if form_type == "AI":
                    co_parts = re.split(r'([/\-])', orig_co)
                    if len(co_parts) >= 9:
                        p4 = re.sub(r'[^A-Za-z0-9]', '', co_parts[6].upper())
                        if p4.endswith("4"):
                            if p4.endswith("A4"): p4 = p4[:-1]
                            else: p4 = p4[:-1] + "A"
                        elif len(p4) > 0 and not p4.endswith("A"):
                            p4 = (p4[:-1] + "A") if p4[-1].isalpha() else (p4 + "A")
                        co_parts[6] = p4
                        
                        p5_raw = "".join(co_parts[8:])
                        p5_clean = re.sub(r'\D.*$', '', p5_raw)[:8]
                        if not p5_clean:
                            p5_clean = p5_raw  
                        orig_co = "".join(co_parts[:8]) + p5_clean

            # Cho phép vượt qua ký tự '|' hoặc '_' ở đầu dòng mới giữa 'Issuance Date:' và '08-APR-2026'
            iss_m = re.search(r'(?i)Issuance\s*Date[\s:;\.\-\|_]*(\d{1,2}-[A-Za-z]{3}-\d{4})', block_fixed)
            orig_date = iss_m.group(1).upper() if iss_m else ""

            # --- ORIGIN CRITERIA ---
            part1, part2 = "", ""
            rvc_m = re.search(r'(RVC\s*[\d\.,]+\s*%?\s*\+?)', block_fixed, re.IGNORECASE)
            if rvc_m: 
                part1 = re.sub(r'\s+', ' ', rvc_m.group(1)).upper().replace(',', '.')

            sub_m = re.search(r'\b(CTSH|CTH|CC)\b', block_fixed, re.IGNORECASE)
            if sub_m: 
                part2 = sub_m.group(1).upper()

            if part1:
                origin = f"{part1} {part2}".strip()
                if origin.endswith('+') and not part2: 
                    origin = origin[:-1].strip()
            elif part2: 
                origin = part2
            else:
                wo_m = re.search(r'\b(W\s*O|WoO|PE|PSR)\b', block_fixed, re.IGNORECASE)
                if wo_m: 
                    val = wo_m.group(1).upper().replace(' ', '')
                    origin = "WO" if val == "WOO" or val == "WO" else val

            # --- DESCRIPTION ---
            desc_zone = re.split(r'(?i)(?:IMPORTING\s+COUNTRY|Original\s+CO|TOTAL|Page)', block_fixed)[0]
            clean_lines = [l.strip() for l in desc_zone.split('\n') if l.strip() and not re.search(r'N/M|\bCARTONS?\b', l, re.I)]
            raw_desc = " ".join(clean_lines)

            # Xóa mã Invoice (nếu còn sót) khỏi Description
            if invoice:
                for inv_code in invoice.replace(" IN", "").split(","):
                    inv_code = inv_code.strip()
                    if len(inv_code) > 2:
                        num_part = inv_code[2:]
                        raw_desc = re.sub(r'\b(?:[VY]N\s*)?' + re.escape(num_part) + r'\b', '', raw_desc, flags=re.IGNORECASE)

            # --- TRÍCH XUẤT BOX 7 QUANTITY VÀ UOM ---
            qty_box7, uom_box7 = extract_box7_qty_scanned(raw_desc, block_fixed)

            if qty_box7 and uom_box7:
                raw_desc = re.sub(r'[-–—]?\s*' + re.escape(qty_box7) + r'[\s\n]*' + re.escape(uom_box7) + r'\b', '', raw_desc, flags=re.IGNORECASE).strip()

            if usd: raw_desc = re.sub(r'(?i)USD\s*' + re.escape(usd), '', raw_desc)
            if date_inv: raw_desc = raw_desc.replace(date_inv, '')
            if qty and uom: 
                raw_desc = re.sub(r'\b' + re.escape(qty) + r'\s+' + re.escape(uom) + r'\b', '', raw_desc, flags=re.IGNORECASE)

            raw_desc = re.sub(r'\b(RVC|W\s*O|WoO|PE|CTSH|CTH|CC|PSR)\b[\d\.\+%\s]*', '', raw_desc, flags=re.IGNORECASE)
            
            raw_desc = re.sub(r'[-–—]\s*(?:USD|EUR|VND)?\s*\d+[\d\.,]*\s*$', '', raw_desc, flags=re.IGNORECASE).strip()
            raw_desc = re.sub(r'^\s*[|\_]+[\s|\_]*(?:[A-Za-z]\s+)?', '', raw_desc).strip()
            raw_desc = re.sub(r'[\s|\_\-\–\—:;\,\.]+$', '', raw_desc).strip()
            raw_desc = re.sub(r'\s+[a-zA-Z]{1,2}$', '', raw_desc).strip()

            desc = clean_text(raw_desc)

            # Box 7 không ghi số lượng -> lấy theo Box 9 (sau khi đã làm sạch mô tả, không ảnh hưởng cột E)
            if not (qty_box7 and uom_box7) and BOX7_FALLBACK_TO_BOX9 and qty and uom:
                qty_box7, uom_box7 = qty, uom
                file_warnings.append(f"Item {item_no}: Box 7 không ghi số lượng → Quantity/UOM - Box 7 lấy theo Box 9 ({qty} {uom})")

            date_inv = normalize_date(date_inv)

            extracted_data.append({
                COLUMNS[0]: form_type, COLUMNS[1]: reference_no, COLUMNS[2]: clean_text(orig_co),
                COLUMNS[3]: item_no, COLUMNS[4]: desc, COLUMNS[5]: clean_text(qty),
                COLUMNS[6]: clean_text(uom), 
                COLUMNS[7]: clean_text(qty_box7), COLUMNS[8]: clean_text(uom_box7),
                COLUMNS[9]: usd, COLUMNS[10]: origin,
                COLUMNS[11]: imp_hs, COLUMNS[12]: exp_hs, COLUMNS[13]: invoice,
                COLUMNS[14]: date_inv, COLUMNS[15]: clean_text(carton), COLUMNS[16]: clean_text(orig_date),
                COLUMNS[17]: "", COLUMNS[18]: date_cert, COLUMNS[19]: exporter,
                COLUMNS[20]: consignee, COLUMNS[21]: transport, COLUMNS[22]: produced_in,
                COLUMNS[23]: exported_to, COLUMNS[24]: "N/M", COLUMNS[25]: box_13_str,
                COLUMNS[26]: third_party_val
            })
            
    logging.info(f"\n" + "="*70)
    logging.info(f"📊 KẾT QUẢ TRÍCH XUẤT (PyMuPDF + Tesseract) FILE: {file_name}")
    logging.info("="*70)
    if not extracted_data:
        print("[!] Không có dữ liệu nào được trích xuất.")
    else:
        for idx, row in enumerate(extracted_data):
            item_id = row.get("Item Number", "N/A")
            logging.info(f"\n📦 --- Dòng hàng hóa thứ {idx + 1} (Item No: {item_id}) ---")
            for key, value in row.items():
                if value and str(value).strip() != "":
                    display_value = str(value)
                    if len(display_value) > 100: display_value = display_value[:97] + "..."
                    logging.info(f"   + {key}: {display_value}")
    print("="*70 + "\n")
    return {"error": None, "data": extracted_data, "file_name": file_name,
            "warnings": file_warnings, "scanned": True}

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
    transport = clean_text(strip_transport_labels(page_first.crop(bbox_transport).extract_text()))
    
    raw_ref = page_first.crop(bbox_ref_no).extract_text()
    ref_match = re.search(r'Reference No\.\s*([A-Z0-9\-]+)', raw_ref, re.IGNORECASE)
    reference_no = ref_match.group(1) if ref_match else clean_text(raw_ref)

    form_type = ""
    try:
        bbox_form = (290, 0, page_first.width, 150)
        cropped_form_page = page_first.crop(bbox_form)
        pil_image = cropped_form_page.to_image(resolution=300).original
        ocr_text = pytesseract.image_to_string(pil_image)
        form_match = re.search(r'FORM\s*([A-Za-z0-9]+)', ocr_text, re.IGNORECASE)
        if form_match: form_type = form_match.group(1).strip().upper()
    except: pass

    movement_cert = ""
    third_party = "No"
    try:
        bbox_box13 = (0, 600, page_last.width, page_last.height)
        cropped_box13 = page_last.crop(bbox_box13)
        img_box13 = cropped_box13.to_image(resolution=300).original
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
    
    # --- BOX 11: Produced in / Exported to theo VỊ TRÍ so với nhãn (không lấy "nước đầu tiên" trong vùng) ---
    # Lỗi cũ: vùng cắt (0,545,...) dính dòng "Third Party Invoicing by: ... 397799 SINGAPORE" ở đáy bảng
    # trang cuối của Form D, nên nước đầu tiên luôn là SINGAPORE và nước thứ hai là nước sản xuất.
    warnings = []
    try:
        produced_in, exported_to, warnings = extract_box11_countries_textlayer(page_last)
    except Exception as e:
        produced_in, exported_to = "", ""
        warnings = [f"Box 11: lỗi khi đọc ({e}) → để trống Produced in / Exported to"]

    if form_type == "E" and not (produced_in and exported_to):
        # Giữ nguyên cách cũ cho Form E khi không định vị được nhãn (Form E đang đúng với cách này)
        asean_china = ["CHINA", "VIETNAM", "MALAYSIA", "SINGAPORE", "INDONESIA", "THAILAND", "PHILIPPINES", "BRUNEI", "CAMBODIA", "LAOS", "MYANMAR"]
        country_matches = re.findall(r'\b(' + '|'.join(asean_china) + r')\b', box11_text, re.IGNORECASE)
        if not produced_in and len(country_matches) > 0: produced_in = country_matches[0].upper()
        if not exported_to and len(country_matches) > 1: exported_to = country_matches[1].upper()
        warnings = []
    
    date_cert = format_to_dd_mm_yyyy(re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', box12_text).group(1)) if re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', box12_text) else ""

    return {
        "exporter": exporter, "consignee": consignee, "transport": transport,
        "reference_no": reference_no, "movement_cert": movement_cert, "third_party": third_party,
        "produced_in": produced_in, "exported_to": exported_to, "date_of_cert": date_cert,
        "form_type": form_type, "warnings": warnings
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
    
    # --- TRÍCH XUẤT BOX 7 QUANTITY VÀ UOM ---
    qty_box7, uom_box7 = "", ""
    box7_match = re.search(r'[-–—]\s*(\d+[\d\.,]*)\s*([A-Za-z]{2,8})\b', desc_cleaned)
    if box7_match:
        qty_box7 = box7_match.group(1).replace(',', '')
        uom_box7 = box7_match.group(2).upper()
    else:
        box7_match_2 = re.search(r'(\d+[\d\.,]*)\s*([A-Za-z]{2,8})\s*(?:\n|\s)*IMPORTING\s+COUNTRY', text, re.IGNORECASE)
        if box7_match_2:
            qty_box7 = box7_match_2.group(1).replace(',', '')
            uom_box7 = box7_match_2.group(2).upper()
        else:
            box7_match_3 = re.search(r'(?:\s|^)(\d+[\d\.,]*)\s*([A-Za-z]{2,8})\s*$', desc_cleaned)
            if box7_match_3:
                qty_box7 = box7_match_3.group(1).replace(',', '')
                uom_box7 = box7_match_3.group(2).upper()
                
    eng_desc = desc_cleaned
    if qty_box7 and uom_box7:
        eng_desc = re.sub(r'[-–—]?\s*' + re.escape(qty_box7) + r'[\s\n]*' + re.escape(uom_box7) + r'\b', '', eng_desc, flags=re.IGNORECASE).strip()
    # ----------------------------------------------
        
    eng_desc = re.sub(r'[-–—:,]\s*$', '', eng_desc).strip()
            
    import_hs = re.sub(r'\D', '', re.search(r'IMPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE).group(1))[:8] if re.search(r'IMPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE) else ""
    export_hs = re.sub(r'\D', '', re.search(r'EXPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE).group(1))[:8] if re.search(r'EXPORTING COUNTRY HS CODE\s*[:\-]?\s*([A-Za-z0-9\.]+)', text, re.IGNORECASE) else ""
    orig_co = re.sub(r'[-–—:.,\s]+$', '', re.search(r'Original CO Reference Number\s*:\s*(.*?)(?=Issuance Date|Issuing Authority|TOTAL|$)', text, re.IGNORECASE).group(1).strip()) if re.search(r'Original CO Reference Number\s*:\s*(.*?)(?=Issuance Date|Issuing Authority|TOTAL|$)', text, re.IGNORECASE) else ""
    issue_date = re.search(r'Issuance Date\s*:\s*(.*?)(?=Issuing Authority|TOTAL|$)', text, re.IGNORECASE).group(1).strip() if re.search(r'Issuance Date\s*:\s*(.*?)(?=Issuing Authority|TOTAL|$)', text, re.IGNORECASE) else ""
    auth = re.search(r'Issuing Authority\s*:\s*(.*?)(?=TOTAL|Page|$)', text, re.IGNORECASE).group(1).strip() if re.search(r'Issuing Authority\s*:\s*(.*?)(?=TOTAL|Page|$)', text, re.IGNORECASE) else ""
    
    return carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth, qty_box7, uom_box7

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
            
            if re.search(r'(?i)(Third\s+Party|Third\s+Country|\bTOTAL\b)', row_text): in_footer_section = True

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
    
    raw_tp = re.split(r'(?i)TOTAL|USD|MYR|EUR', clean_text(third_party_text))[0].strip()
    clean_tp = re.sub(r'(?i)^(Third\s+Party\s+Invoicing\s+by\s*[:;]?|Third\s+Party\s*[:;]?|Third\s+Country\s*[:;]?)', '', raw_tp).strip()
    return items, global_invoice, clean_tp

def process_single_pdf(file_data):
    file_name = file_data["name"]
    file_bytes = file_data["bytes"]
    extracted_data = []
    
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            if not first_page_text or len(first_page_text.strip()) < 50:
                return process_scanned_pdf(pdf, file_bytes, file_name)

            global_info = extract_global_info(pdf.pages[0], pdf.pages[-1])
            items, global_invoice, third_party_column_val = extract_table_items(pdf)
            box_13_str = "YES" if (global_info["third_party"] == "Yes" and global_info["movement_cert"] != "") else "No"
            
            if global_info["form_type"] == "AI":
                third_party_column_val = "DESIPRO PTE LTD 230 STADIUM BOULEVARD 397799 SINGAPORE"

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
                        invoice = invoice.replace(date_match.group(1), "")
                    
                    # --- XỬ LÝ NHIỀU INVOICE, BỎ KHOẢNG TRẮNG BỊ LỖI DO OCR ---
                    inv_matches = re.findall(r'[VY]N\s*\d+(?:\s+\d+)*(?:\-\d+)?', invoice, re.IGNORECASE)
                    if not inv_matches:
                        inv_matches = re.findall(r'[VY]N[A-Z0-9\-]+', invoice, re.IGNORECASE)
                        
                    if inv_matches:
                        cleaned_invs = []
                        for m in inv_matches:
                            cln = re.sub(r'\s+', '', m).upper().replace('YN', 'VN')
                            if cln not in cleaned_invs:
                                cleaned_invs.append(cln)
                        invoice_number = ",".join(cleaned_invs)
                    else:
                        invoice_number = re.sub(r'\s+', '', invoice).strip(',;-')
                
                carton, eng_desc, qty, uom, import_hs, export_hs, orig_co, issue_date, auth, qty_box7, uom_box7 = parse_description_fields(desc, weight_value_text)
                
                usd_match = re.search(r'USD\s*([\d,\.]+)', weight_value_text, re.IGNORECASE)
                usd = usd_match.group(1).strip() if usd_match else ""
                if not usd:
                    usd_match_desc = re.search(r'USD\s*([\d,\.]+)', desc, re.IGNORECASE)
                    usd = usd_match_desc.group(1).strip() if usd_match_desc else ""
                
                item_no_val = clean_text(item["item_no"])
                if item_no_val.upper() == "CONTINUATION": item_no_val = ""
                
                extracted_data.append({
                    COLUMNS[0]: global_info["form_type"], 
                    COLUMNS[1]: global_info["reference_no"],
                    COLUMNS[2]: clean_text(orig_co), 
                    COLUMNS[3]: item_no_val, 
                    COLUMNS[4]: clean_text(eng_desc),
                    COLUMNS[5]: clean_text(qty), 
                    COLUMNS[6]: clean_text(uom), 
                    COLUMNS[7]: clean_text(qty_box7),
                    COLUMNS[8]: clean_text(uom_box7),
                    COLUMNS[9]: usd, 
                    COLUMNS[10]: clean_text(item["origin"]), 
                    COLUMNS[11]: import_hs, 
                    COLUMNS[12]: export_hs,       
                    COLUMNS[13]: invoice_number, 
                    COLUMNS[14]: invoice_date, 
                    COLUMNS[15]: clean_text(carton),
                    COLUMNS[16]: format_to_dd_mm_yyyy(issue_date), 
                    COLUMNS[17]: clean_text(auth),
                    COLUMNS[18]: global_info["date_of_cert"], 
                    COLUMNS[19]: global_info["exporter"],
                    COLUMNS[20]: global_info["consignee"], 
                    COLUMNS[21]: global_info["transport"],
                    COLUMNS[22]: global_info["produced_in"], 
                    COLUMNS[23]: global_info["exported_to"],
                    COLUMNS[24]: clean_text(item["marks"]), 
                    COLUMNS[25]: box_13_str,
                    COLUMNS[26]: third_party_column_val
                })
    except Exception as e:
        return {"error": f"{file_name}: Lỗi trích xuất - {str(e)}", "data": [], "file_name": file_name,
                "warnings": [], "scanned": False}
    return {"error": None, "data": extracted_data, "file_name": file_name,
            "warnings": global_info.get("warnings", []), "scanned": False}

# ==========================================
# 4. GIAO DIỆN STREAMLIT (chỉ trình bày & luồng tương tác – không chứa logic trích xuất)
# ==========================================
APP_CSS = """
<style>
:root {
    --coo-accent: #006BA6;          /* chữ nhấn trên nền sáng: tương phản 5.6:1 */
    --coo-btn: #0073AE;             /* nền nút chính, chữ trắng: tương phản 5.0:1 */
    --coo-btn-hover: #005E8F;
    --coo-muted: #5B6472;           /* chữ phụ trên nền sáng: 6.1:1 */
    --coo-border: rgba(15, 23, 42, 0.10);
    --coo-shadow: 0 1px 2px rgba(15, 23, 42, 0.05), 0 6px 16px rgba(15, 23, 42, 0.05);
    --coo-radius: 14px;
    --coo-dur: 180ms;
    --coo-ease: cubic-bezier(0.2, 0.8, 0.2, 1);
    --coo-hover: rgba(15, 23, 42, 0.04);
    /* Màu trạng thái (nền nhạt + chữ đậm, tương phản >= 6:1) */
    --coo-ok: #188038;       --coo-ok-fg: #FFFFFF;  --coo-ok-bg: #E6F4EA;   --coo-ok-text: #0D652D;  --coo-ok-line: rgba(24, 128, 56, 0.28);
    --coo-warn-bg: #FEF3D6;  --coo-warn-text: #7A4B00;
    --coo-err-bg: #FCE8E6;   --coo-err-text: #A50E0E;
    --coo-new-bg: rgba(15, 23, 42, 0.07); --coo-new-text: #47505E;
}
@media (prefers-color-scheme: dark) {
    :root {
        --coo-accent: #5CC0F2;      /* trên nền tối #0E1117: 9.1:1 */
        --coo-muted: #A3ACB9;       /* trên nền tối: 8.2:1 */
        --coo-border: rgba(255, 255, 255, 0.12);
        --coo-shadow: 0 1px 2px rgba(0, 0, 0, 0.35);
        --coo-hover: rgba(255, 255, 255, 0.05);
        --coo-ok: #34C759;       --coo-ok-fg: #06250F;  --coo-ok-bg: rgba(52, 199, 89, 0.13);  --coo-ok-text: #86E3A1;  --coo-ok-line: rgba(52, 199, 89, 0.38);
        --coo-warn-bg: rgba(255, 204, 0, 0.16);  --coo-warn-text: #FFD666;
        --coo-err-bg: rgba(255, 69, 58, 0.18);   --coo-err-text: #FF9F97;
        --coo-new-bg: rgba(255, 255, 255, 0.10); --coo-new-text: #C3CAD4;
    }
}
#MainMenu, footer, [data-testid="stAppDeployButton"] { visibility: hidden; }
.block-container { padding-top: 4rem; max-width: 1400px; }

.coo-header { margin: 0 0 1.5rem 0; animation: coo-fade 240ms var(--coo-ease); }
.coo-title { font-size: 2rem; font-weight: 750; letter-spacing: -0.025em; line-height: 1.2; margin: 0; }
.coo-title .hl { color: var(--coo-accent); }

/* --- Thanh tổng số file đã tải lên --- */
.coo-count { display: flex; align-items: center; gap: 0.8rem; padding: 0.75rem 0.95rem; margin: 0.15rem 0 0.55rem 0;
             border-radius: 12px; background: var(--coo-ok-bg); border: 1px solid var(--coo-ok-line);
             animation: coo-pop 260ms var(--coo-ease); }
.coo-count-icon { flex: none; width: 30px; height: 30px; border-radius: 50%; display: grid; place-items: center;
                  background: var(--coo-ok); color: var(--coo-ok-fg); font-weight: 800; font-size: 0.95rem; }
.coo-count-main { display: flex; align-items: baseline; gap: 0.4rem; flex-wrap: wrap; min-width: 0; }
.coo-count-num { font-size: 1.5rem; font-weight: 750; line-height: 1; letter-spacing: -0.02em; font-variant-numeric: tabular-nums; }
.coo-count-label { font-weight: 600; font-size: 0.95rem; }
.coo-count-meta { margin-left: auto; flex: none; color: var(--coo-muted); font-size: 0.86rem; font-variant-numeric: tabular-nums; }
.coo-count.empty { background: var(--coo-hover); border: 1px dashed var(--coo-border); animation: none; }
.coo-count.empty .coo-count-icon { background: var(--coo-new-bg); color: var(--coo-muted); }
.coo-count.empty .coo-count-label { font-weight: 600; }
.coo-count.empty .coo-count-meta { margin-left: 0; display: block; margin-top: 0.1rem; }

/* --- Danh sách tệp đã tải lên (cuộn trong khung, hợp với hàng trăm file) --- */
.coo-files { max-height: 276px; overflow-y: auto; border: 1px solid var(--coo-border); border-radius: 12px;
             scrollbar-width: thin; scroll-behavior: smooth; overscroll-behavior: contain;
             animation: coo-fade 220ms var(--coo-ease); }
.coo-files:focus-visible { outline: 2px solid var(--coo-btn); outline-offset: 2px; }
.coo-file { display: grid; grid-template-columns: 2.2rem minmax(0, 1fr) auto auto; align-items: center; column-gap: 0.65rem;
            padding: 0.5rem 0.8rem; font-size: 0.9rem; line-height: 1.35; border-bottom: 1px solid var(--coo-border);
            transition: background-color var(--coo-dur) ease; }
.coo-file:last-child { border-bottom: 0; }
.coo-file:hover { background: var(--coo-hover); }
.coo-file[style] { animation: coo-fade 260ms var(--coo-ease) both; animation-delay: calc(var(--i) * 22ms); }
.coo-file-idx { text-align: right; color: var(--coo-muted); font-size: 0.78rem; font-variant-numeric: tabular-nums; }
.coo-file-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 500; }
.coo-file-tags { display: inline-flex; gap: 0.3rem; }
.coo-file-size { min-width: 4.2rem; text-align: right; color: var(--coo-muted); font-size: 0.8rem; font-variant-numeric: tabular-nums; }
.coo-tag { font-size: 0.72rem; font-weight: 650; padding: 0.14rem 0.55rem; border-radius: 999px; white-space: nowrap; }
.coo-tag.ok { background: var(--coo-ok-bg); color: var(--coo-ok-text); }
.coo-tag.warn { background: var(--coo-warn-bg); color: var(--coo-warn-text); }
.coo-tag.err { background: var(--coo-err-bg); color: var(--coo-err-text); }
.coo-tag.new { background: var(--coo-new-bg); color: var(--coo-new-text); }
.coo-note { display: flex; gap: 0.5rem; align-items: flex-start; margin-top: 0.6rem; padding: 0.55rem 0.75rem; border-radius: 10px;
            background: var(--coo-warn-bg); color: var(--coo-warn-text); font-size: 0.86rem; }
@media (max-width: 640px) {
    .coo-file { grid-template-columns: 1.8rem minmax(0, 1fr) auto; }
    .coo-file-size { display: none; }
    .coo-count-meta { display: none; }
}

.coo-step { font-size: 0.78rem; font-weight: 650; letter-spacing: 0.06em; text-transform: uppercase;
            color: var(--coo-muted); margin-bottom: 0.35rem; }
.coo-hint { color: var(--coo-muted); font-size: 0.86rem; }

/* Thẻ (container border=True): bo góc + đổ bóng nhẹ */
/* (Nếu Streamlit đổi tên lớp, thẻ chỉ mất bóng, không vỡ bố cục.) */
.st-key-coo_upload, .st-key-coo_action, .st-key-coo_results {
    border-radius: var(--coo-radius) !important; box-shadow: var(--coo-shadow);
    animation: coo-fade var(--coo-dur) ease-out;
}

/* Nút: nút chính màu thương hiệu, hiệu ứng hover/press nhẹ */
button[kind="primary"], [data-testid="stBaseButton-primary"] {
    background-color: var(--coo-btn) !important; border-color: var(--coo-btn) !important; color: #FFFFFF !important;
}
button[kind="primary"]:hover:not(:disabled), [data-testid="stBaseButton-primary"]:hover:not(:disabled) {
    background-color: var(--coo-btn-hover) !important; border-color: var(--coo-btn-hover) !important;
}
button[kind="primary"]:disabled, [data-testid="stBaseButton-primary"]:disabled {
    opacity: 0.45; cursor: not-allowed;
}
.stButton button, .stDownloadButton button {
    border-radius: 10px !important; font-weight: 600 !important; min-height: 44px;
    transition: transform var(--coo-dur) ease, box-shadow var(--coo-dur) ease, background-color var(--coo-dur) ease;
}
.stButton button:hover:not(:disabled), .stDownloadButton button:hover:not(:disabled) {
    transform: translateY(-1px); box-shadow: var(--coo-shadow);
}
.stButton button:active:not(:disabled), .stDownloadButton button:active:not(:disabled) { transform: translateY(0) scale(0.98); }

[data-testid="stMetric"] { border: 1px solid var(--coo-border); border-radius: 12px; padding: 0.7rem 0.9rem;
                           animation: coo-fade var(--coo-dur) ease-out; }
[data-testid="stMetricLabel"] p { color: var(--coo-muted); }

/* Khung xổ danh sách tệp: bo góc đồng bộ với thẻ */
.st-key-coo_upload [data-testid="stExpander"] details { border-radius: 12px; }

@keyframes coo-fade { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: none; } }
@keyframes coo-pop { from { opacity: 0; transform: scale(0.98); } to { opacity: 1; transform: none; } }
@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after { animation: none !important; transition: none !important; }
}
</style>
"""

STATUS_OK, STATUS_WARN, STATUS_ERR = "✅ Thành công", "⚠️ Cảnh báo", "❌ Lỗi"
NUMERIC_EXPORT_COLS = ["Item Number", "Quantity - Box 9", "Quantity - Box 7", "USD",
                       "IMPORTING COUNTRY HS CODE", "EXPORTING COUNTRY HS CODE", "CARTON"]


def init_session_state():
    defaults = {
        "session_id": str(uuid.uuid4()),
        "has_recorded_access": False,
        "pdf_files": [],
        "is_processing": False,
        "extracted_data": None,
        "errors": [],
        "file_status": [],
        "excel_bytes": None,
        "run_seconds": 0.0,
        "cancelled": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_data_state():
    st.session_state.extracted_data = None
    st.session_state.errors = []
    st.session_state.file_status = []
    st.session_state.excel_bytes = None
    st.session_state.run_seconds = 0.0


def build_excel_bytes(rows):
    """Giữ nguyên cách xuất Excel cũ (tên cột, thứ tự, kiểu số của các cột số)."""
    df_result = pd.DataFrame(rows, columns=COLUMNS)
    for col in NUMERIC_EXPORT_COLS:
        if col in df_result.columns:
            df_result[col] = df_result[col].astype(str).str.replace(',', '')
            df_result[col] = pd.to_numeric(df_result[col], errors='coerce')
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False, sheet_name="C_O_FormE")
    return output.getvalue()


def friendly_error(raw_error):
    """Đổi thông báo kỹ thuật thành lý do ngắn + hướng xử lý."""
    msg = str(raw_error or "")
    low = msg.lower()
    if "tesseract" in low and ("not installed" in low or "not in your path" in low or "notfound" in low):
        return "Máy chủ chưa có Tesseract OCR → cài Tesseract hoặc đặt biến môi trường TESSERACT_CMD rồi chạy lại."
    if "password" in low or "encrypt" in low:
        return "File PDF có mật khẩu → mở khóa file rồi tải lên lại."
    if "no /root" in low or "eof" in low or "pdfsyntax" in low or "cannot open" in low or "broken" in low:
        return "File PDF bị hỏng hoặc tải lên chưa trọn → mở và lưu lại bằng trình đọc PDF rồi thử lại."
    if "index out of range" in low or "nonetype" in low:
        return "Bố cục không giống mẫu C/O Form D/E/AI → kiểm tra đúng loại chứng từ; nếu đúng, gửi file cho bộ phận kỹ thuật."
    detail = msg.split("Lỗi trích xuất - ", 1)[-1]
    return f"Không đọc được file ({detail[:120]}) → kiểm tra file có đúng mẫu C/O và không bị scan quá mờ."


def summarize_result(result):
    """Chuyển kết quả của process_single_pdf thành 1 dòng trạng thái cho bảng tiến độ."""
    warnings = result.get("warnings") or []
    if result.get("error"):
        status, note = STATUS_ERR, friendly_error(result["error"])
    elif not result.get("data"):
        status, note = STATUS_ERR, "Không tìm thấy dòng hàng nào → kiểm tra file có đúng mẫu C/O."
    elif warnings:
        status = STATUS_WARN
        note = warnings[0] if len(warnings) == 1 else f"{warnings[0]} (+{len(warnings) - 1} cảnh báo khác)"
    else:
        status, note = STATUS_OK, ""
    return {
        "Tệp": result.get("file_name", ""),
        "Loại": "Scan (OCR)" if result.get("scanned") else "Text",
        "Trạng thái": status,
        "Số dòng": len(result.get("data") or []),
        "Ghi chú": note,
        "Chi tiết cảnh báo": "\n".join(warnings),
    }


def status_frame(rows, sort=True):
    df = pd.DataFrame(rows, columns=["Tệp", "Loại", "Trạng thái", "Số dòng", "Ghi chú", "Chi tiết cảnh báo"])
    order = {STATUS_ERR: 0, STATUS_WARN: 1, STATUS_OK: 2}
    if sort and not df.empty:
        df = df.sort_values("Trạng thái", key=lambda s: s.map(order), kind="stable").reset_index(drop=True)
    return df


STATUS_COLUMN_CONFIG = {
    "Tệp": st.column_config.TextColumn("Tệp", width="medium"),
    "Loại": st.column_config.TextColumn("Loại", width="small"),
    "Trạng thái": st.column_config.TextColumn("Trạng thái", width="small"),
    "Số dòng": st.column_config.NumberColumn("Số dòng", width="small"),
    "Ghi chú": st.column_config.TextColumn("Ghi chú (lý do & hướng xử lý)", width="large"),
    "Chi tiết cảnh báo": None,   # ẩn trên bảng, vẫn có trong dữ liệu
}


def render_header():
    st.markdown(APP_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="coo-header">
            <p class="coo-title"><span class="hl">PDF</span> Extract to <span class="hl">Excel</span> Pipeline</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


FILE_TAGS = {STATUS_OK: ("ok", "Thành công"), STATUS_WARN: ("warn", "Cảnh báo"), STATUS_ERR: ("err", "Lỗi")}
FILE_SEARCH_MIN = 8          # từ bao nhiêu file thì hiện ô tìm kiếm
FILE_ANIMATED_ROWS = 12      # chỉ các dòng đầu có hiệu ứng lần lượt, để danh sách hàng trăm file vẫn nhẹ


def _format_size(num_bytes):
    num_bytes = num_bytes or 0
    if num_bytes >= 1024 * 1024: return f"{num_bytes / (1024 * 1024):.1f} MB"
    if num_bytes >= 1024: return f"{num_bytes / 1024:.0f} KB"
    return f"{num_bytes} B"


def _file_list_html(shown, name_counts, status_by_name, has_results):
    """Danh sách tệp dạng HTML một khối (một dòng, không thụt lề) – nhanh kể cả với hàng trăm file."""
    rows = []
    for pos, (idx, f) in enumerate(shown):
        name = html_escape(f.name)
        tags = ""
        status = status_by_name.get(f.name)
        if status in FILE_TAGS:
            css_class, label = FILE_TAGS[status]
            tags += f'<span class="coo-tag {css_class}">{label}</span>'
        elif has_results:
            tags += '<span class="coo-tag new">Chưa xử lý</span>'
        if name_counts[f.name] > 1:
            tags += '<span class="coo-tag warn">Trùng tên</span>'
        stagger = f' style="--i:{pos}"' if pos < FILE_ANIMATED_ROWS else ""
        rows.append(
            f'<div class="coo-file" role="listitem"{stagger}>'
            f'<span class="coo-file-idx">{idx}</span>'
            f'<span class="coo-file-name" title="{name}">{name}</span>'
            f'<span class="coo-file-tags">{tags}</span>'
            f'<span class="coo-file-size">{_format_size(getattr(f, "size", 0))}</span>'
            f'</div>'
        )
    return ('<div class="coo-files" role="list" tabindex="0" aria-label="Danh sách tệp đã tải lên">'
            + "".join(rows) + '</div>')


def render_uploaded_files(files):
    """Cho người dùng thấy ngay: ĐÃ TẢI LÊN BAO NHIÊU file và ĐÓ LÀ NHỮNG FILE NÀO (tên đầy đủ, tìm được)."""
    count = len(files)
    if not count:
        st.markdown(
            '<div class="coo-count empty"><span class="coo-count-icon">↑</span>'
            '<span><span class="coo-count-label">Chưa có file nào được tải lên</span>'
            '<span class="coo-count-meta">Kéo thả hoặc bấm Upload · chọn nhiều file cùng lúc bằng Ctrl/Shift</span></span></div>',
            unsafe_allow_html=True,
        )
        return

    total_bytes = sum(getattr(f, "size", 0) or 0 for f in files)
    name_counts = Counter(f.name for f in files)
    dup_names = [name for name, n in name_counts.items() if n > 1]
    st.markdown(
        '<div class="coo-count" role="status"><span class="coo-count-icon">✓</span>'
        f'<span class="coo-count-main"><span class="coo-count-num">{count}</span>'
        '<span class="coo-count-label">file đã tải lên</span></span>'
        f'<span class="coo-count-meta">Tổng dung lượng {_format_size(total_bytes)}</span></div>',
        unsafe_allow_html=True,
    )

    status_by_name = {s["Tệp"]: s["Trạng thái"] for s in st.session_state.file_status}
    has_results = st.session_state.extracted_data is not None
    with st.expander(f"Danh sách tệp đã tải lên ({count})", expanded=True):
        query = ""
        if count > FILE_SEARCH_MIN:
            query = st.text_input(
                "Tìm tệp", key="coo_file_search", placeholder="Tìm theo tên tệp, rồi nhấn Enter…",
                label_visibility="collapsed", disabled=st.session_state.is_processing,
            ).strip().lower()
        shown = [(idx, f) for idx, f in enumerate(files, start=1) if query in f.name.lower()]
        if query:
            st.markdown(f'<div class="coo-hint">Khớp {len(shown)}/{count} tệp với “{html_escape(query)}”</div>',
                        unsafe_allow_html=True)
        if shown:
            st.markdown(_file_list_html(shown, name_counts, status_by_name, has_results), unsafe_allow_html=True)
        else:
            st.markdown('<div class="coo-hint">Không có tệp nào khớp. Xóa từ khóa để xem lại toàn bộ danh sách.</div>',
                        unsafe_allow_html=True)

    if dup_names:
        preview = ", ".join(html_escape(n) for n in dup_names[:3]) + ("…" if len(dup_names) > 3 else "")
        st.markdown(
            f'<div class="coo-note"><span>⚠</span><span>{len(dup_names)} tên tệp bị tải lên trùng ({preview}). '
            'Mỗi bản sẽ được trích xuất một lần nên Excel sẽ có dòng lặp – hãy bỏ bớt bản trùng nếu không cần.</span></div>',
            unsafe_allow_html=True,
        )


def render_upload_card():
    with st.container(border=True, key="coo_upload"):
        st.markdown('<div class="coo-step">Bước 1 · Nguồn dữ liệu PDF</div>', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "Kéo thả File vào đây", type=["pdf"], accept_multiple_files=True,
            disabled=st.session_state.is_processing,
        )
        st.session_state.pdf_files = uploaded_files or []
        render_uploaded_files(st.session_state.pdf_files)


def render_action_card():
    """Hành động chính duy nhất: Bắt đầu trích xuất (trở thành nút phụ khi đã có kết quả)."""
    has_files = len(st.session_state.pdf_files) > 0
    has_results = st.session_state.extracted_data is not None
    with st.container(border=True, key="coo_action"):
        st.markdown('<div class="coo-step">Bước 2 · Trích xuất</div>', unsafe_allow_html=True)
        label = "Trích xuất lại" if has_results else "🚀 BẮT ĐẦU TRÍCH XUẤT"
        start = st.button(
            label, type="secondary" if has_results else "primary", width="stretch",
            disabled=not has_files or st.session_state.is_processing,
        )
        if not has_files:
            st.markdown('<div class="coo-hint">Tải PDF lên ở Bước 1 để bắt đầu.</div>', unsafe_allow_html=True)
        elif has_results:
            st.markdown('<div class="coo-hint">Đã có kết quả bên dưới. Bấm “Trích xuất lại” nếu bạn đổi danh sách file.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="coo-hint">Sẵn sàng xử lý {len(st.session_state.pdf_files)} file. '
                        'File scan dùng OCR nên chậm hơn file có text.</div>', unsafe_allow_html=True)
    if start:
        record_file_processing(st.session_state.session_id, len(st.session_state.pdf_files))
        reset_data_state()
        st.session_state.cancelled = False
        st.session_state.is_processing = True
        st.rerun()


def run_extraction():
    """Chạy trích xuất từng file, hiển thị tiến độ + trạng thái từng file theo thời gian thực."""
    files = [{"name": f.name, "bytes": f.getvalue()} for f in st.session_state.pdf_files]
    total = len(files)
    all_rows, errors, statuses = [], [], []
    started = time.time()

    with st.status(f"Đang trích xuất 0/{total} file…", expanded=True) as status_box:
        st.button("Dừng", key="cancel_run", on_click=_cancel_run)
        progress = st.progress(0.0)
        current = st.empty()
        live_table = st.empty()
        for idx, f_data in enumerate(files, start=1):
            current.markdown(f'<div class="coo-hint">Đang xử lý: <b>{f_data["name"]}</b></div>', unsafe_allow_html=True)
            result = process_single_pdf(f_data)
            if result["error"]:
                errors.append(result["error"])
            else:
                all_rows.extend(result["data"])
            statuses.append(summarize_result(result))

            elapsed = time.time() - started
            eta = elapsed / idx * (total - idx)
            progress.progress(idx / total, text=f"{idx}/{total} file · còn khoảng {int(eta // 60)} phút {int(eta % 60)} giây")
            status_box.update(label=f"Đang trích xuất {idx}/{total} file…", expanded=True)
            live_table.dataframe(status_frame(statuses[-8:][::-1], sort=False), hide_index=True, width="stretch",
                                 column_config=STATUS_COLUMN_CONFIG)
            del result
            gc.collect()
        current.empty()
        status_box.update(label=f"Hoàn tất {total} file trong {time.time() - started:.0f} giây", state="complete",
                          expanded=False)

    st.session_state.extracted_data = all_rows
    st.session_state.errors = errors
    st.session_state.file_status = statuses
    st.session_state.excel_bytes = build_excel_bytes(all_rows) if all_rows else None
    st.session_state.run_seconds = time.time() - started
    st.session_state.is_processing = False
    st.rerun()


def _cancel_run():
    # Bấm "Dừng" khiến Streamlit ngắt lượt chạy hiện tại; callback này chạy trước lượt chạy mới.
    st.session_state.is_processing = False
    st.session_state.cancelled = True
    reset_data_state()


def render_results():
    rows = st.session_state.extracted_data
    statuses = st.session_state.file_status
    n_ok = sum(1 for s in statuses if s["Trạng thái"] == STATUS_OK)
    n_warn = sum(1 for s in statuses if s["Trạng thái"] == STATUS_WARN)
    n_err = sum(1 for s in statuses if s["Trạng thái"] == STATUS_ERR)

    with st.container(border=True, key="coo_results"):
        st.markdown('<div class="coo-step">Bước 3 · Kết quả & xuất file</div>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 1.6], vertical_alignment="bottom")
        c1.metric("Thành công", n_ok)
        c2.metric("Cảnh báo", n_warn)
        c3.metric("Lỗi", n_err)
        c4.metric("Dòng dữ liệu", len(rows))
        with c5:
            if st.session_state.excel_bytes:
                st.download_button(
                    label="📥 TẢI FILE EXCEL (.XLSX)", data=st.session_state.excel_bytes,
                    file_name="Decathlon_Data.xlsx", type="primary", width="stretch",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.button("📥 TẢI FILE EXCEL", disabled=True, width="stretch")
            st.markdown(f'<div class="coo-hint">Xử lý {len(statuses)} file trong {st.session_state.run_seconds:.0f} giây.</div>',
                        unsafe_allow_html=True)

        if n_err:
            st.error(f"{n_err} file không trích xuất được. Xem cột “Ghi chú” trong tab Trạng thái để biết lý do và cách xử lý.")
        elif n_warn:
            st.warning(f"{n_warn} file có ô cần kiểm tra lại (để trống hoặc lấy giá trị thay thế). Xem tab Trạng thái.")

        status_tab = f"Trạng thái từng file ({len(statuses)})"
        data_tab = f"Dữ liệu chi tiết ({len(rows)} dòng)"
        tab_names = [status_tab, data_tab] if (n_err or n_warn) else [data_tab, status_tab]
        tabs = dict(zip(tab_names, st.tabs(tab_names)))
        with tabs[status_tab]:
            df_status = status_frame(statuses)
            st.dataframe(df_status, hide_index=True, width="stretch", column_config=STATUS_COLUMN_CONFIG,
                         height=min(420, 38 + 35 * max(len(df_status), 1)))
            warn_rows = [s for s in statuses if s["Chi tiết cảnh báo"]]
            if warn_rows:
                with st.expander("Xem toàn bộ cảnh báo"):
                    for s in warn_rows:
                        st.markdown(f"**{s['Tệp']}**")
                        st.text(s["Chi tiết cảnh báo"])
        with tabs[data_tab]:
            if rows:
                st.dataframe(pd.DataFrame(rows, columns=COLUMNS), width="stretch", height=450)
            else:
                st.info("Chưa có dòng dữ liệu nào. Kiểm tra tab Trạng thái để biết lý do.")


def main():
    st.set_page_config(page_title="PDF Extract to Excel", layout="wide", page_icon="📑")
    init_session_state()

    # Ghi nhận ngay khi truy cập web lần đầu trong phiên
    if not st.session_state.has_recorded_access:
        record_access(st.session_state.session_id)
        st.session_state.has_recorded_access = True

    render_header()
    left, right = st.columns([3, 2], gap="large")
    with left:
        render_upload_card()
    with right:
        render_action_card()

    if st.session_state.cancelled:
        st.info("Đã dừng trích xuất. Kết quả dở dang đã được bỏ qua – bấm “Bắt đầu trích xuất” để chạy lại.")
        st.session_state.cancelled = False

    if st.session_state.is_processing:
        run_extraction()
    elif st.session_state.extracted_data is not None:
        render_results()


if __name__ == "__main__":
    main()
