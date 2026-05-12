# File: src/utils.py
import io
import re
import time
import json
import os
import pdfplumber
import httpx
from groq import Groq
from dotenv import load_dotenv

# --- 0. TẢI BIẾN MÔI TRƯỜNG ---
# Hàm này sẽ tự động tìm file .env trong thư mục gốc và tải các biến vào hệ thống
load_dotenv()

# --- 1. CẤU HÌNH API KEY & CỘT DỮ LIỆU ---
# Lấy API Key từ file .env (An toàn, không còn hard-code)
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

# Cảnh báo sớm nếu quên cấu hình file .env
if not GROQ_API_KEY:
    raise ValueError("❌ LỖI: Không tìm thấy GROQ_API_KEY. Vui lòng kiểm tra lại file .env!")

COLUMNS = [
    "Products consigned from (Exporter's business name, address, country)",
    "Products consigned to (Consignee's name, address, country)",
    "Means of transport and route (as far as known)",
    "Reference No",
    "Item Number",
    "Marks and numbers on packages",
    "Number and type of packages, description of products (including quantity where appropriate and HS number in six digit code)",
    "Origin criteria (see Overleaf Notes)",
    "Gross weight or net weight or other quantity, and value (FOB) only when RVC criterion is applied",
    "Number, date of Invoices",
    "Movement Certificate",
    "Third Party Invoicing",
]

# --- 2. PROMPTS CHO LLM ---
HEADER_PROMPT = """
You are extracting global information from PAGE 1 of an ASEAN-China Form E.
Extract the exact text for the exporter, consignee, transport details, and reference number.
Also, find the global Invoice Number and Date (usually in Box 10, combining both lines like "VN26000069 CN 19/04/2026").
Return ONLY JSON:
{
  "exporter": "[Full text of Box 1]",
  "consignee": "[Full text of Box 2]",
  "transport": "[Full text of Box 3]",
  "reference_no": "[Reference Number at top right]",
  "global_invoice": "[Combine Invoice Number and Date]"
}
"""

ITEM_CHUNK_PROMPT = """
You are an elite Data Engineer parsing a SPATIALLY-FORMATTED Form E.
The text preserves physical spaces. Visually, the columns from left to right are:
[Item No] | [Marks] | [Description (Box 7)] | [Origin Criteria (Box 8)] | [Weight/Value] | [Invoice & Date (Box 10)]

CRITICAL RULES:
1. ALIGNMENT: Read vertically based on spacing! Box 10 (Invoice) is always on the FAR RIGHT. Box 8 (Origin like 'PE', 'CTH') is in the middle.
2. BOX 10 (INVOICES): Often spans 2 lines (Number above, Date below). Combine them! (e.g., "VN26000069 CN 19/04/2026").
3. CONTINUATION TEXT: Text like "Original CO Reference...", "Issuing Authority:", or long HS Codes will drop to the next lines WITHOUT an item number. Extract this as a separate JSON object with "item_number": "CONTINUATION" so my system can merge it.
4. KEEP BOTH "IMPORTING COUNTRY HS CODE" and "EXPORTING COUNTRY HS CODE" in the description.

Return ONLY a JSON array. Example format:
[
  {
    "item_number": "1",
    "marks": "N/M",
    "description": "1 CARTON Goalpost - 18 PCE / IMPORTING HS 950699 / EXPORTING HS 950699",
    "origin_criteria": "PE",
    "quantity_value": "18 PCE / USD 203.58",
    "invoice": "VN26000069 CN 19/04/2026"
  },
  {
    "item_number": "CONTINUATION",
    "marks": "",
    "description": "Original CO Reference Number: E12345 / Issuance Date: 01-JAN",
    "origin_criteria": "",
    "quantity_value": "",
    "invoice": ""
  }
]
"""

# --- 3. CORE FUNCTIONS ---
def extract_pdf_by_chunks(pdf_bytes: bytes):
    pages_text = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                text = page.extract_text(layout=True)
            except:
                text = page.extract_text(x_tolerance=2, keep_blank_chars=True) # Fallback
                
            if text:
                pages_text.append(text)
            
    page_1_text = pages_text[0] if pages_text else ""
    full_pdf_text = "\n".join(pages_text)
    
    chunks = []
    CHUNK_SIZE = 3 # Gửi 3 trang/lần để tránh vượt Rate Limit
    for i in range(0, len(pages_text), CHUNK_SIZE):
        chunk_str = "\n\n--- NEXT PAGE ---\n\n".join(pages_text[i:i + CHUNK_SIZE])
        chunks.append(chunk_str)
        
    return page_1_text, full_pdf_text, chunks


def call_llm(prompt: str, content: str, is_array=False):
    # Dùng httpx chống lỗi kết nối mạng (timeout)
    custom_http_client = httpx.Client(http2=False, verify=False, timeout=120.0)
    client = Groq(api_key=GROQ_API_KEY, http_client=custom_http_client, max_retries=3)
    
    for attempt in range(4):
        try:
            res = client.chat.completions.create(
                messages=[{"role": "system", "content": prompt}, {"role": "user", "content": content}],
                model="llama-3.3-70b-versatile",
                temperature=0,
            )
            text = res.choices[0].message.content.strip()
            pattern = r'\[.*\]' if is_array else r'\{.*\}'
            match = re.search(pattern, text, re.DOTALL)
            
            if match:
                return json.loads(match.group(0))
            else:
                time.sleep(3)
                continue
        except Exception as e:
            err_str = str(e).lower()
            if "429" in err_str or "limit" in err_str or "connection" in err_str or "timeout" in err_str:
                time.sleep(15) 
            else:
                time.sleep(5)
            continue
    return [] if is_array else {}
