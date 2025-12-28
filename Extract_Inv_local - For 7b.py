import os
import sys
import requests
import json
import re
import pandas as pd
import base64
import io
import gc
import time
import logging
from pypdf import PdfReader
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance

# --- 1. การตั้งค่า (Configuration) ---
# เปลี่ยนเป็น /api/generate เพื่อความเสถียรของ Vision Model บน Ollama 
OLLAMA_API_URL = "http://127.0.0.1:11434/api/generate"
MODEL_NAME = "scb10x/typhoon-ocr1.5-3b:latest" 
#MODEL_NAME = "scb10x/typhoon-ocr-7b:latest" 
POPPLER_PATH = r"C:\poppler\Library\bin"

# ตั้งค่า Path สำหรับ Source และ Output
if len(sys.argv) >= 3:
    SOURCE_DIR, OUTPUT_DIR, PAGE_CONFIG = sys.argv[1], sys.argv[2], sys.argv[3]
else:
    SOURCE_DIR = r"d:\Project\ocr\source"
    OUTPUT_DIR = r"d:\Project\ocr\output"
    PAGE_CONFIG = "all" # หรือระบุหน้า เช่น "1,2,5"

if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)

# --- 2. การตั้งค่า Logging ---
logging.basicConfig(
    filename=os.path.join(OUTPUT_DIR, "ocr_process_log.txt"),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)
console = logging.StreamHandler()
logging.getLogger('').addHandler(console)

# --- 3. ฟังก์ชันสนับสนุน (Helper Functions) ---
def check_ollama_connection():
    try:
        # ตรวจสอบว่า Service Ollama เปิดอยู่หรือไม่
        response = requests.get("http://127.0.0.1:11434/api/tags", timeout=5)
        return response.status_code == 200
    except:
        return False

def preprocess_image(image, max_size=1536):
    """ปรับแต่งภาพให้ AI อ่านง่ายขึ้น"""
    if image.mode != 'RGB': image = image.convert('RGB')
    image = ImageEnhance.Contrast(image).enhance(1.5) # เพิ่มความคมชัด 
    if max(image.size) > max_size: 
        image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
    return image

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

# --- 4. ฟังก์ชัน OCR หลัก (ปรับปรุง Payload สำหรับ /api/generate) ---
def extract_text_from_image(file_path, pages_list, filename, output_dir):
    extracted_data = []
    poppler = POPPLER_PATH if os.path.exists(POPPLER_PATH) else None
    
    for page_num in pages_list:
        try:
            logging.info(f"--- Processing: {filename} (Page {page_num}) ---")
            
            # แปลง PDF เป็น Image
            images = convert_from_path(file_path, first_page=page_num, last_page=page_num, poppler_path=poppler, dpi=300)
            if not images: continue

            img = preprocess_image(images[0])
            base_name = os.path.splitext(filename)[0]
            
            # บันทึกภาพ PNG เพื่อตรวจสอบ
            img_path = os.path.join(output_dir, f"{base_name}_page{page_num}.png")
            img.save(img_path, format="PNG")
            
            # แปลงภาพเป็น Base64
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")

            # เตรียม Payload สำหรับ /api/generate (ลดโอกาส AI ทวนคำสั่ง) [cite: 38]
            payload = {
                "model": MODEL_NAME,
                # ใช้ Prompt ที่สั้นและเน้นผลลัพธ์ที่เป็นข้อความเท่านั้น
                "prompt": "OCR the following image and output the text content accurately in Markdown format. Keep the original language (Thai/English). No explanations.",
                "images": [img_str],
                "stream": False,
                "options": {
                    "temperature": 0.0,
                    "num_ctx": 4096,
                    "top_p": 0.9,
                    "repeat_penalty": 1.1
                }       
            }

            start_ai = time.time()
            response = requests.post(OLLAMA_API_URL, json=payload, timeout=300)
            
            if response.status_code == 200:
                # ดึงข้อมูลจากฟิลด์ 'response' (สำหรับ /api/generate)
                raw_content = response.json().get("response", "").strip()
                logging.info(f"   [Timer] AI Time: {time.time() - start_ai:.2f}s")
                
                if len(raw_content) > 10:
                    # บันทึกไฟล์ TXT ทันที [cite: 53]
                    txt_name = f"{base_name}_page{page_num}.txt"
                    with open(os.path.join(output_dir, txt_name), 'w', encoding='utf-8') as f:
                        f.write(raw_content)
                    logging.info(f"   [Success] Saved: {txt_name}")
                    extracted_data.append((page_num, raw_content))
                else:
                    logging.warning(f"   [Warning] Page {page_num} returned insufficient text.")
            else:
                logging.error(f"   [Error] Ollama API Error {response.status_code}: {response.text}")

            del img, img_str, buffered, images
            gc.collect()
        except Exception as e:
            logging.error(f"   [Fatal Error] Page {page_num}: {str(e)}")
            
    return extracted_data

# --- 5. ฟังก์ชันสกัดข้อมูลด้วย Regex (Data Parsing) ---
def parse_ocr_data(text):
    if not text: return "", "", "", "", ""
    # ค้นหาเลขที่ใบกำกับ/ใบแจ้งหนี้
    inv = re.search(r"(?:เลขที่|No\.?)\s*[:\.]?\s*([A-Za-z0-9\-\/]{4,})", text, re.IGNORECASE)
    # ค้นหาวันที่
    dt = re.search(r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", text)
    # ค้นหายอดเงินสุทธิ
    amt = re.search(r"(?:ยอดสุทธิ|Total|รวมทั้งสิ้น)\s*[:\.]?\s*([\d,]+\.\d{2})", text)
    # ค้นหาเลขประจำตัวผู้เสียภาษี (13 หลัก)
    tax_ids = re.findall(r"\b(\d{13})\b", text)
    
    return (
        inv.group(1).strip() if inv else "",
        dt.group(1).strip() if dt else "",
        amt.group(1) if amt else "",
        tax_ids[0] if tax_ids else "",
        "00000" if "สำนักงานใหญ่" in text else ""
    )

# --- 6. ฟังก์ชันหลัก (Main Process) ---
def main():
    logging.info("================ OCR Job Started ================")
    if not check_ollama_connection():
        logging.error("Cannot connect to Ollama. Please ensure Ollama is running.")
        return
    
    data_rows = []
    files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".pdf")]
    
    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        try:
            reader = PdfReader(file_path)
            target_pages = get_target_pages(PAGE_CONFIG, len(reader.pages))
            
            ocr_results = extract_text_from_image(file_path, target_pages, filename, OUTPUT_DIR)
                
            for p_num, raw_text in ocr_results:
                inv, dt, amt, tx, br = parse_ocr_data(raw_text)
                data_rows.append({
                    "FileName": filename,
                    "Page": p_num,
                    "Tax_ID": tx,
                    "Invoice_No": inv,
                    "Date": dt,
                    "Amount": amt
                })
        except Exception as e:
            logging.error(f"Error processing {filename}: {e}")

    # บันทึกผลลัพธ์ลง Excel [cite: 28]
    if data_rows:
        df = pd.DataFrame(data_rows)
        excel_path = os.path.join(OUTPUT_DIR, "summary_ocr_results.xlsx")
        df.to_excel(excel_path, index=False)
        logging.info(f"Summary Excel created: {excel_path}")
    
    logging.info("================ OCR Job Finished ================")

if __name__ == "__main__":
    main()