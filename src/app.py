import streamlit as st
import pandas as pd
import io
from concurrent.futures import ThreadPoolExecutor
from processor import process_single_file

# --- UI CONFIGURATION ---
st.set_page_config(page_title="Decathlon | PDF Data Intelligence", layout="wide")

# Decathlon Blue: #0082C3
DECATHLON_BLUE = "#0082C3"

st.markdown(f"""
    <style>
    .stButton>button {{
        background-color: {DECATHLON_BLUE};
        color: white;
        border-radius: 4px;
        font-weight: bold;
        border: none;
    }}
    .main {{ background-color: #F8FAFC; }}
    </style>
    """, unsafe_allow_html=True)

def main():
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/12/Decathlon_Logo.svg/1280px-Decathlon_Logo.svg.png", width=200)
    st.title("Data Extraction Portal")
    st.info("Efficiently process Form E documents with Llama-3 AI acceleration.")

    # 1. UPLOAD AREA (Drag & Drop)
    uploaded_files = st.file_uploader(
        "Upload PDF Folder / Files", 
        type=["pdf"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("START BATCH PROCESSING", use_container_width=True):
            all_data = []
            progress_bar = st.progress(0)
            status_area = st.empty()
            
            # 2. MULTI-THREADING IMPLEMENTATION
            # Max_workers controls how many files are processed simultaneously
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = {executor.submit(process_single_file, f): f.name for f in uploaded_files}
                
                for i, future in enumerate(concurrent.futures.as_completed(futures)):
                    f_name = futures[future]
                    try:
                        result = future.result()
                        all_data.extend(result)
                        
                        # Update Progress
                        percent = (i + 1) / len(uploaded_files)
                        progress_bar.progress(percent)
                        status_area.write(f"✅ Processed: **{f_name}**")
                    except Exception as e:
                        st.error(f"Error in {f_name}: {e}")

            # 3. RESULT DISPLAY & EXPORT
            if all_data:
                df = pd.DataFrame(all_data)
                st.subheader("Extracted Data Preview")
                
                # Modern Table with Search/Filter
                st.dataframe(df, use_container_width=True)
                
                # CTA: Excel Export
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False)
                
                st.download_button(
                    label="EXPORT TO EXCEL (.XLSX)",
                    data=output.getvalue(),
                    file_name="Decathlon_Extraction.xlsx",
                    mime="application/vnd.ms-excel"
                )

if __name__ == "__main__":
    main()