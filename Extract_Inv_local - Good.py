import os
import sys
import requests
import json
import re
import pandas as pd
import base64
import io
import gc
from pypdf import PdfReader
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance

# --- ส่วนการตั้งค่า (Configuration) ---
OLLAMA_API_URL = "http://localhost:11434/api/generate" 
MODEL_NAME = "scb10x/typhoon-ocr1.5-3b:latest" 
POPPLER_PATH = r"C:\poppler\Library\bin"

if len(sys.argv) >= 3:
    SOURCE_DIR, OUTPUT_DIR, PAGE_CONFIG = sys.argv[1], sys.argv[2], sys.argv[3]
else:
    SOURCE_DIR, OUTPUT_DIR, PAGE_CONFIG = r"d:\Project\ocr\source", r"d:\Project\ocr\output", "2"

os.makedirs(OUTPUT_DIR, exist_ok=True)

def check_ollama_connection():
    try:
        response = requests.get("http://localhost:11434/api/tags", timeout=5)
        return response.status_code == 200
    except: return False

def load_vendor_master():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Vendor_branch.xlsx")
    if not os.path.exists(path): return None
    try:
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        df['เลขประจำตัวผู้เสียภาษี'] = df['เลขประจำตัวผู้เสียภาษี'].fillna('').str.replace(r'\D', '', regex=True)
        df['สาขา'] = df['สาขา'].fillna('').str.strip()
        return df
    except: return None

def get_target_pages(selection_str, total_pages):
    selection_str = str(selection_str).lower().replace(" ", "")
    if selection_str == 'all': return list(range(1, total_pages + 1))
    pages = set()
    for part in selection_str.split(','):
        if '-' in part:
            try:
                s, e = part.split('-')
                pages.update(range(int(s), (total_pages if e == 'n' else int(e)) + 1))
            except: pass
        elif part.isdigit(): pages.add(int(part))
    return sorted([p for p in pages if 1 <= p <= total_pages])

def preprocess_image(image, max_size=1280):
    if image.mode != 'RGB': image = image.convert('RGB')
    image = ImageEnhance.Contrast(image).enhance(1.8)
    if max(image.size) > max_size: image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
    return image

def extract_text_from_image(file_path, pages_list):
    extracted_pages = []
    poppler = POPPLER_PATH if os.path.exists(POPPLER_PATH) else None
    for page_num in pages_list:
        try:
            print(f"   [Step 1] Rendering Page {page_num}...")
            images = convert_from_path(file_path, first_page=page_num, last_page=page_num, poppler_path=poppler, dpi=300)
            if not images: continue
            img = preprocess_image(images[0])
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
            print(f"   [Step 2] Sending to AI...")
            payload = {
                "model": MODEL_NAME,
                "prompt": "Extract text from image. Return clean Markdown only.",
                "images": [img_str],
                "stream": False,
                "options": {"temperature": 0, "num_ctx": 4096}
            }
            response = requests.post(OLLAMA_API_URL, json=payload, timeout=300)
            if response.status_code == 200:
                raw_content = response.json().get("response", "").strip()
                if "Instructions:" in raw_content:
                    raw_content = raw_content.split("Instructions:")[-1]
                cleaned = clean_ocr_text(raw_content)
                extracted_pages.append((page_num, cleaned))
                print(f"   [Step 3] Page {page_num} Processed.")
            del img, img_str, buffered, images
            gc.collect()
        except Exception as e: print(f"   [Error] Page {page_num}: {e}")
    return extracted_pages

def clean_ocr_text(text):
    if not text: return ""
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n', '\n\n', text)
    return text.strip()

def parse_ocr_data(text):
    if not text: return "", "", "", "", ""
    inv = re.search(r"(?:เลขที่|No\.?)\s*[:\.]?\s*([A-Za-z0-9\-\/]{4,})", text, re.IGNORECASE)
    invoice_no = inv.group(1).strip() if inv else ""
    dt = re.search(r"(?:วันที่|Date)\s*[:\.]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", text)
    bill_date = dt.group(1).strip() if dt else ""
    all_amounts = re.findall(r"([\d,]+\.\d{2})", text)
    amount = all_amounts[-1] if all_amounts else ""
    tax_ids = re.findall(r"\b(\d{13})\b", text)
    tax_id = tax_ids[0] if tax_ids else ""
    branch = "00000" if re.search(r"(สำนักงานใหญ่|สนญ|Head\s*Office|HQ)", text, re.IGNORECASE) else ""
    if not branch:
        br_m = re.search(r"(?:สาขา|Branch)(?:\s*ที่|\s*No\.?)?\s*[:\.]?\s*(\d{1,5})", text)
        branch = br_m.group(1).zfill(5) if br_m else ""
    return invoice_no, bill_date, amount, tax_id, branch

def extract_additional_fields(text):
    res = {'description': '', 'sales_promotion': '', 'total_amount': '', 'withholding_tax': ''}
    if not text: return res
    # Regex ดึง Description โดยหาจุดเริ่ม "จำนวนเงิน 1" และจุดจบยอดเงินทศนิยม
    desc_match = re.search(r"จำนวนเงิน\s+1\s+(.+?)(?=\s+[\d,]+\.\d{2})", text, re.DOTALL)
    if desc_match:
        res['description'] = desc_match.group(1).strip()
    else:
        fallback = re.search(r"1\s+([\u0E00-\u0E7F].+?)(?=\s+[\d,]+\.\d{2})", text, re.DOTALL)
        if fallback: res['description'] = fallback.group(1).strip()

    sales = re.search(r'ค่าส่งเสริมการขาย[^\n]+', text)
    if sales: res['sales_promotion'] = sales.group(0).strip()
    total = re.search(r'รวมภาษีมูลค่าเพิ่ม\s*[:\.]?\s*([\d,]+\.\d{2})', text)
    if total: res['total_amount'] = total.group(1)
    wht = re.search(r'หัก\s*ณ\s*ที่จ่าย\s+([\d,]+\.\d{2})', text)
    if wht: res['withholding_tax'] = wht.group(1)
    return res

def main():
    print(f"--- OCR Processing (Column Optimized Mode) ---")
    if not check_ollama_connection(): return
    vendor_df = load_vendor_master()
    data_rows = []
    files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".pdf")]
    
    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        print(f"\n[File] {filename}")
        try:
            reader = PdfReader(file_path)
            ocr_results = extract_text_from_image(file_path, get_target_pages(PAGE_CONFIG, len(reader.pages)))
            for p_num, raw_text in ocr_results:
                txt_path = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_page{p_num}.txt")
                with open(txt_path, 'w', encoding='utf-8') as f: f.write(raw_text)
                
                inv, dt, amt, tx, br = parse_ocr_data(raw_text)
                add = extract_additional_fields(raw_text)
                
                data_rows.append({
                    "Link PDF": f'=HYPERLINK("{file_path}", "{filename}")',
                    "Page": p_num,
                    "VendorID_OCR": tx,
                    "Branch_OCR": br,
                    "Invoice No": inv,
                    "Date": dt,
                    "Amount": amt,
                    "Total Amount": add['total_amount'],
                    "Withholding Tax": add['withholding_tax'],
                    "Description": add['description'],
                    "Sales Promotion": add['sales_promotion'],
                    "สาขาในใบเสร็จ": br
                })
        except Exception as e: print(f"   [Error] {filename}: {e}")

    if data_rows:
        df = pd.DataFrame(data_rows)
        if vendor_df is not None:
            df = pd.merge(df, vendor_df, left_on=['VendorID_OCR', 'Branch_OCR'], right_on=['เลขประจำตัวผู้เสียภาษี', 'สาขา'], how='left')
        
        # จัดเรียงคอลัมน์และลบ "เลขประจำตัวผู้เสียภาษี_y" ออกตามคำขอ
        cols_order = [
            "Link PDF", "Page", "VendorID_OCR", "Branch_OCR", "Vendor code SAP", 
            "ชื่อบริษัท", "Invoice No", "Date", "Amount", "Total Amount", 
            "Withholding Tax", "Description", "Sales Promotion", "สาขา", "สาขาในใบเสร็จ"
        ]
        
        df = df.reindex(columns=[c for c in cols_order if c in df.columns])
        
        excel_path = os.path.join(OUTPUT_DIR, "summary_ocr_local_final.xlsx")
        df.to_excel(excel_path, index=False)
        print(f"\n[Success] Created Excel: {excel_path}")

if __name__ == "__main__": main()