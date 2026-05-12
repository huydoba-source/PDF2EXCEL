import io
import time
import pandas as pd
import concurrent.futures
from typing import List
from utils import extract_pdf_by_chunks, call_llm, HEADER_PROMPT, ITEM_CHUNK_PROMPT, COLUMNS

def process_single_file(file_obj):
    """
    Handles a single file's extraction logic. 
    This is the unit of work for our thread pool.
    """
    file_name = file_obj.name
    pdf_bytes = file_obj.read()
    
    # 1. Text Extraction
    page_1_text, full_pdf_text, chunks = extract_pdf_by_chunks(pdf_bytes)
    
    # 2. Header Data
    headers = call_llm(HEADER_PROMPT, page_1_text, is_array=False)
    global_invoice = headers.get("global_invoice", "")
    
    # 3. Logic-based flags
    full_text_upper = full_pdf_text.upper()
    m_status = "Yes" if any(x in full_text_upper for x in ["ORIGINAL CO", "MOVEMENT"]) else "No"
    t_status = "Yes" if "THIRD PARTY" in full_text_upper else "No"

    # 4. Item extraction (with internal chunk loop)
    raw_items = []
    for chunk in chunks:
        items = call_llm(ITEM_CHUNK_PROMPT, chunk, is_array=True)
        raw_items.extend(items)
    
    # 5. Data Flattening
    file_results = []
    for item in raw_items:
        # (Merging logic simplified for brevity - same as your original)
        file_results.append({
            COLUMNS[0]: headers.get("exporter", ""),
            COLUMNS[3]: headers.get("reference_no", ""),
            COLUMNS[4]: item.get("item_number", ""),
            COLUMNS[6]: item.get("description", ""),
            COLUMNS[9]: item.get("invoice", global_invoice),
            COLUMNS[10]: m_status,
            COLUMNS[11]: t_status,
        })
    return file_results
