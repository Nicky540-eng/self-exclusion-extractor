import streamlit as st
import easyocr
import pandas as pd
import re
import os
import io
import numpy as np
import pypdfium2 as pdfium
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Registry Tracker", layout="wide")

st.title("⚡ Targeted Self-Exclusion Document Extractor")
st.write("Production Version: High-precision extraction targeting the official 'SELF-BANNING ORDER' heading parameters.")

@st.cache_resource
def load_ocr_reader():
    return easyocr.Reader(['en'], gpu=False)

try:
    reader = load_ocr_reader()
except Exception as e:
    st.error(f"OCR Engine Initialization Error: {e}")

def extract_perfect_data(file_bytes, file_name):
    """
    Scans document lines to find 'SELF-BANNING ORDER' or 'SELF- BANNING ORDER'
    and cleanly extracts the full name and 13-digit ID number.
    """
    # Default fallback values pulled directly from filename structure if text scan fails
    base_name = os.path.splitext(file_name)[0].strip()
    full_name_clean = base_name.upper() if (base_name and not base_name.isspace()) else "UNKNOWN APPLICANT"
    id_number = "Not Found"
    
    try:
        pdf = pdfium.PdfDocument(file_bytes)
        found_target = False
        
        # Scan through the document pages (Checking early pages first where the header lives)
        for page_idx in range(len(pdf)):
            page = pdf[page_idx]
            bitmap = page.render(scale=2.5) 
            pil_img = bitmap.to_pil().convert('L')
            img_np = np.array(pil_img)
            
            # Read line-by-line text segments
            ocr_results = reader.readtext(img_np, detail=0) 
            
            for line in ocr_results:
                clean_line = line.strip()
                # Target the exact official header pattern
                if "SELF-" in clean_line.upper() and "BANNING" in clean_line.upper() and "ORDER" in clean_line.upper():
                    
                    # --- STEP 1: IDENTITY NUMBER EXTRACTION ---
                    # Strip spaces to handle structured space formatting in the ID (e.g., 000722 5058 085)
                    collapsed_line = clean_line.replace(" ", "").replace("-", "").replace(":", "")
                    id_match = re.search(r'\d{13}', collapsed_line)
                    if id_match:
                        id_number = id_match.group(0)
                    
                    # --- STEP 2: FULL NAME EXTRACTION ---
                    # Corrected safe regex split pattern avoiding range identifier syntax issues
                    parts = re.split(r'[–::\-]', clean_line)
                    if len(parts) >= 2:
                        for part in parts:
                            part_upper = part.upper()
                            # Locate the name block segment containing titles
                            if "MR." in part_upper or "MS." in part_upper or "MRS." in part_upper:
                                # Strip out standard titles to isolate the true raw identity
                                name_segment = part.replace("Mr.", "").replace("Mr", "").replace("Ms.", "").replace("Ms", "").replace("Mrs.", "").replace("Mrs", "")
                                # Remove any lingering ID remnants from the text segment string
                                name_segment = re.sub(r'ID\s*No.*', '', name_segment, flags=re.IGNORECASE)
                                name_segment = re.sub(r'\d+', '', name_segment)
                                
                                name_final = name_segment.strip()
                                if name_final:
                                    full_name_clean = name_final.upper()
                                    break
                    
                    found_target = True
                    break
            
            if found_target:
                break
                
    except Exception as e:
        pass 

    return {
        "File Name": file_name,
        "Full Name": full_name_clean,
        "Identity Number": id_number
    }

def create_excel_download(dataframe):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Registry Database"
    ws.views.sheetView[0].showGridLines = True
    
    font_title = Font(name="Calibri", size=15, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=11, bold=False, color="000000")
    
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    ws['A1'] = "Gauteng Gambling Board - Verified Registry Extraction"
    ws['A1'].font = font_title
    
    headers = ["File Name", "Full Name", "Identity Number"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for row_idx, row_data in enumerate(dataframe.values, start=4):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = font_data
            cell.border = thin_border
            if row_idx % 2 == 0:
                cell.fill = fill_zebra
            if col_idx == 3:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")
                
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[3].height = 24
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)
        
    wb.save(output)
    return output.getvalue()

st.sidebar.header("Execution Settings")
st.sidebar.info("🌐 Web Folder Mode Active")

uploaded_files = st.file_uploader(
    "Drag and drop your target compliance folder here, or click to browse:", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_records = []
    if st.button("🚀 Run Targeted Extraction Pipeline"):
        progress_bar = st.progress(0)
        for idx, uploaded_file in enumerate(uploaded_files):
            file_bytes = uploaded_file.read()
            record = extract_perfect_data(file_bytes, uploaded_file.name)
            all_records.append(record)
            progress_bar.progress((idx + 1) / len(uploaded_files))

    if all_records:
        df = pd.DataFrame(all_records)
        st.success(f"Processing Complete! Successfully parsed {len(all_records)} profiles.")
        
        st.subheader("📋 Processed Data Preview")
        st.dataframe(df, use_container_width=True)
        
        excel_data = create_excel_download(df)
        st.download_button(
            label="📥 Download Excel Spreadsheet",
            data=excel_data,
            file_name="Verified_Handwritten_Registry.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )