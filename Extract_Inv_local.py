import os
import sys
import requests
import json
import re
import pandas as pd
import base64
import io
import gc
import platform
import shutil
import subprocess
import time
from pypdf import PdfReader
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance

# --- Cross-platform Configuration ---
def get_default_poppler_path():
    """Get Poppler path based on operating system"""
    system = platform.system()
    
    if system == 'Windows':
        # Common Windows install locations
        possible_paths = [
            r"C:\poppler\Library\bin",
            r"C:\poppler\bin",
            r"C:\Program Files\poppler\bin",
            os.path.expanduser(r"~\poppler\bin"),
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None  # Will try system PATH
    
    elif system == 'Darwin':  # macOS
        possible_paths = [
            "/opt/homebrew/bin",
            "/usr/local/bin",
            "/opt/homebrew/Cellar/poppler",
        ]
        for path in possible_paths:
            if os.path.exists(path) and shutil.which("pdftoppm"):
                return None  # Use system PATH
        return None
    
    else:  # Linux (WSL included)
        # Check if poppler-utils is installed
        if shutil.which("pdftoppm"):
            return None  # Use system PATH
        return None

def get_default_source_dir():
    """Get default source directory based on OS"""
    system = platform.system()
    if system == 'Windows':
        return r"d:\Project\ocr\source"
    else:
        # Use home directory for Linux/Mac/WSL
        return os.path.expanduser("~/ocr/source")

def get_default_output_dir():
    """Get default output directory based on OS"""
    system = platform.system()
    if system == 'Windows':
        return r"d:\Project\ocr\output"
    else:
        # Use home directory for Linux/Mac/WSL
        return os.path.expanduser("~/ocr/output")

# --- vLLM Configuration (Updated for Typhoon OCR 1.5 2B) ---
# Point to the vLLM server running on localhost:8000
VLLM_API_URL = os.environ.get("VLLM_API_URL", "http://localhost:8000/v1/chat/completions")
MODEL_NAME = "typhoon-ai/typhoon-ocr1.5-2b" 
POPPLER_PATH = os.environ.get("POPPLER_PATH", get_default_poppler_path())

# Script directory for relative paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_FILE = "document_templates.json"

# Command line arguments
if len(sys.argv) >= 3:
    SOURCE_DIR, OUTPUT_DIR, PAGE_CONFIG = sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "All"
    DOC_TYPE = sys.argv[4] if len(sys.argv) > 4 else "auto"
else:
    SOURCE_DIR = get_default_source_dir()
    OUTPUT_DIR = get_default_output_dir()
    PAGE_CONFIG = "2"
    DOC_TYPE = "auto"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Load Document Templates ---
def load_templates():
    """Load document templates from JSON file"""
    path = os.path.join(SCRIPT_DIR, TEMPLATES_FILE)
    if not os.path.exists(path):
        print(f"Warning: Templates file not found: {TEMPLATES_FILE}")
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading templates: {e}")
        return None

def detect_document_type(text, templates):
    """Auto-detect document type based on keywords in text"""
    if not text or not templates:
        return "invoice"
    
    text_lower = text.lower()
    priority_types = ["cy_instruction"]
    
    for priority_type in priority_types:
        template = templates.get("templates", {}).get(priority_type, {})
        keywords = template.get("detect_keywords", [])
        for keyword in keywords:
            if keyword.lower() in text_lower:
                return priority_type
    
    lines = text.split('\n')
    header_text = "\n".join(lines[:15]).lower()
    scores = {}
    
    for doc_type, template in templates.get("templates", {}).items():
        if doc_type in priority_types:
            continue
        keywords = template.get("detect_keywords", [])
        score = 0
        for keyword in keywords:
            kw_low = keyword.lower()
            if kw_low in header_text:
                score += 10
            if kw_low in text_lower:
                score += 1
        if score > 0:
            scores[doc_type] = score
    
    if scores:
        return max(scores, key=scores.get)
    return "invoice"

def extract_field_by_patterns(text, patterns, options=None):
    """Extract field value using multiple regex patterns"""
    if not text or not patterns:
        return ""
    options = options or {}
    for pattern in patterns:
        try:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                value = match.group(1) if match.lastindex and match.lastindex >= 1 else match.group(0)
                if options.get("clean_html"):
                    value = re.sub(r'<br\s*/?>', ' ', value)
                    value = re.sub(r'<[^>]+>', '', value)
                value = re.sub(r'[\r\n]+', ' ', value)
                value = re.sub(r'\s+', ' ', value).strip()
                if options.get("min_digits"):
                    digit_count = len(re.sub(r'\D', '', value))
                    if digit_count < options["min_digits"]:
                        continue
                if options.get("clean_non_digits"):
                    value = re.sub(r'\D', '', value)
                    if options.get("length"):
                        value = value[:options["length"]]
                if value:
                    return value
        except Exception:
            continue
    return ""

def extract_common_fields(text, common_fields_config, doc_type_name=None):
    """Extract common fields (tax_id, branch)"""
    result = {"tax_id": "", "branch": ""}
    if not text or not common_fields_config:
        return result
    
    tax_config = common_fields_config.get("tax_id", {})
    tax_patterns = tax_config.get("patterns", [])
    COMPANY_TAX_ID = "0105522018355"
    
    SIAM_CONTAINER_TAX_ID = "0105531101901"
    SAHATHAI_TAX_ID = "0107560000192"
    MON_LOGISTICS_TAX_ID = "0105559135291"
    
    if "สยามคอนเทนเนอร์ เทอร์มินอล" in text or "สยามคอนเทนเนอร์เทอร์มินอล" in text:
        result["tax_id"] = SIAM_CONTAINER_TAX_ID
    elif "สหไทย เทอร์มินอล" in text or "สหไทยเทอร์มินอล" in text:
        result["tax_id"] = SAHATHAI_TAX_ID
    elif "มนต์โลจิสติกส์ เซอร์วิส" in text or "มนต์โลจิสติกส์เซอร์วิส" in text:
        result["tax_id"] = MON_LOGISTICS_TAX_ID
    else:
        all_tax_ids = re.findall(r"\b(\d{13})\b", text)
        vendor_tax_ids = [tid for tid in all_tax_ids if tid != COMPANY_TAX_ID]
        if vendor_tax_ids:
            result["tax_id"] = vendor_tax_ids[0]
        else:
            all_dashed = re.findall(r"\b(\d{1}-\d{4}-\d{5}-\d{2}-\d{1})\b", text)
            for match in all_dashed:
                clean_id = re.sub(r"\D", "", match)
                if clean_id != COMPANY_TAX_ID:
                    result["tax_id"] = clean_id
                    break
            if not result["tax_id"]:
                spaced_matches = re.findall(r"\b(\d{1}\s+\d{12})\b", text)
                for match in spaced_matches:
                    clean_id = re.sub(r"\D", "", match)
                    if clean_id != COMPANY_TAX_ID:
                        result["tax_id"] = clean_id
                        break
            if not result["tax_id"]:
                for pattern in tax_patterns:
                    value = extract_field_by_patterns(text, [pattern], {"clean_non_digits": True, "length": 13})
                    if value and len(value) >= 10 and value != COMPANY_TAX_ID:
                        result["tax_id"] = value
                        break
    
    branch_config = common_fields_config.get("branch", {})
    default_hq = branch_config.get("default_hq", "00000")
    pad_zeros = branch_config.get("pad_zeros", 5)
    
    checked_branch_match = re.search(r'[☑✓✔]\s*สาขา(?:ที่)?\s*(\d+)', text)
    if checked_branch_match:
        result["branch"] = checked_branch_match.group(1).zfill(pad_zeros)
    else:
        checked_hq_match = re.search(r'[☑✓✔]\s*(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office)', text, re.IGNORECASE)
        if checked_hq_match:
            result["branch"] = default_hq
        else:
            vendor_branch_match = re.search(r'สาขาที่ออกใบกำกับภาษี\s*(?:คือ|:)?\s*(\d{1,5})', text, re.IGNORECASE)
            if vendor_branch_match:
                result["branch"] = vendor_branch_match.group(1).zfill(pad_zeros)
            else:
                hq_with_num_match = re.search(r'(?:สำนักงานใหญ่|HEAD\s*OFFICE)\s*[:\s]?\s*(\d{5})', text, re.IGNORECASE)
                if hq_with_num_match:
                    result["branch"] = hq_with_num_match.group(1).zfill(pad_zeros)
                else:
                    branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})(?!\d)", text, re.IGNORECASE)
                    if branch_match:
                        result["branch"] = branch_match.group(1).zfill(pad_zeros)
                    else:
                        ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
                        if ho_match:
                            result["branch"] = default_hq
                        else:
                            result["branch"] = default_hq
    return result

def parse_ocr_data_with_template(text, templates, doc_type="auto"):
    """Parse OCR text using document template patterns"""
    result = {
        "document_type": "", "document_type_name": "", "document_no": "",
        "date": "", "amount": "", "tax_id": "", "branch": "", "extra_fields": {}
    }
    
    if not text:
        return result
    if not templates:
        templates = load_templates()
    if not templates:
        return parse_ocr_data_basic(text)
    
    if doc_type == "auto":
        detected_type = detect_document_type(text, templates)
    else:
        detected_type = doc_type if doc_type in templates.get("templates", {}) else "invoice"
    
    result["document_type"] = detected_type
    template = templates.get("templates", {}).get(detected_type, {})
    result["document_type_name"] = template.get("name", detected_type)
    
    common_fields = templates.get("common_fields", {})
    common_result = extract_common_fields(text, common_fields, result["document_type_name"])
    result["tax_id"] = common_result["tax_id"]
    result["branch"] = common_result["branch"]
    
    fields_config = template.get("fields", {})
    for field_name, field_config in fields_config.items():
        patterns = field_config.get("patterns", [])
        options = {
            "clean_html": field_config.get("clean_html", False),
            "clean_non_digits": field_config.get("clean_non_digits", False),
            "length": field_config.get("length"),
            "min_digits": field_config.get("min_digits")
        }
        text_to_search = text
        skip_lines = field_config.get("skip_lines", 0)
        if skip_lines > 0:
            lines = text.split('\n')
            text_to_search = '\n'.join(lines[skip_lines:])
        
        value = extract_field_by_patterns(text_to_search, patterns, options)
        
        if not value and field_config.get("fallback") == "last_amount":
            amounts = re.findall(r"([\d,]+\.\d{2})", text)
            value = amounts[-1] if amounts else ""
        
        if field_name == "document_no" and value:
            digits_only = re.sub(r'\D', '', value)
            if len(digits_only) == 13 and digits_only.isdigit():
                value = ""
        
        if field_name == "date" and not value:
            date_patterns = [
                r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})',
                r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2})'
            ]
            for date_pattern in date_patterns:
                date_matches = re.findall(date_pattern, text)
                if date_matches:
                    value = date_matches[0]
                    break
        if field_name == "date" and value:
            value = re.sub(r'[-.]', '/', value)
        
        if field_name in ["document_no", "date", "amount"]:
            result[field_name] = value
        else:
            result["extra_fields"][field_name] = value
    
    return result

def parse_ocr_data_basic(text):
    """Basic OCR parsing without templates"""
    result = {
        "document_type": "invoice", "document_type_name": "ใบกำกับภาษี/Invoice",
        "document_no": "", "date": "", "amount": "", "tax_id": "", "branch": "", "extra_fields": {}
    }
    if not text:
        return result
    
    # Updated: Support 'เอกสารเลขที่', 'Document No.', 'Ref. Invoice No.'
    inv_match = re.search(r"(?:เลขที่|เอกสารเลขที่|Document\s*No\.?|Ref\.\s*Invoice\s*No\.?)\s*[:\.]?\s*([A-Za-z0-9\-\/]{3,})", text, re.IGNORECASE)
    result["document_no"] = inv_match.group(1) if inv_match else ""
    
    # Updated: Support 'วัน เดือน ปี', 'Date'
    date_match = re.search(r"(?:วันที่|วัน\s*เดือน\s*ปี|Date)\s*[:\.]?\s*(\d{1,2}\s+[^\s]+\s+\d{4}|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", text, re.IGNORECASE)
    result["date"] = date_match.group(1) if date_match else ""
    
    amounts = re.findall(r"([\d,]+\.\d{2})", text)
    result["amount"] = amounts[-1] if amounts else ""
    
    all_tax_ids = re.findall(r"\b(\d{13})\b", text)
    if all_tax_ids:
        result["tax_id"] = all_tax_ids[0]
    else:
        tax_pattern_match = re.search(r"\b\d{1}-\d{4}-\d{5}-\d{2}-\d{1}\b", text)
        if tax_pattern_match:
            result["tax_id"] = re.sub(r"\D", "", tax_pattern_match.group(0))
    
    ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
    if ho_match:
        result["branch"] = "00000"
    else:
        branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})", text, re.IGNORECASE)
        if branch_match:
            result["branch"] = branch_match.group(1).zfill(5)
    return result

# --- Updated Check Connection Function for vLLM ---
def check_vllm_connection():
    """Check if vLLM is running and accessible"""
    try:
        # Check vLLM models endpoint
        url = VLLM_API_URL.replace("/chat/completions", "/models")
        response = requests.get(url, timeout=5)
        return response.status_code == 200
    except:
        return False

# --- Auto-Start vLLM Functions ---
def start_vllm_server():
    """Start vLLM server automatically"""
    print("[System] Starting vLLM server...")
    
    # Construct command using sys.executable to ensure we use the same Python environment
    cmd = [
        sys.executable, "-m", "vllm.entrypoints.openai.api_server",
        "--model", "typhoon-ai/typhoon-ocr1.5-2b",
        "--trust-remote-code",
        "--dtype", "half", # Using half as per user's previous successful run
        "--max-model-len", "4096",
        "--device", "cuda",
        "--port", "8000"
    ]
    
    try:
        # Start process detached, user can see output in console
        process = subprocess.Popen(cmd)
        return process
    except Exception as e:
        print(f"[Error] Failed to start vLLM: {e}")
        return None

def ensure_vllm_running():
    """Check if vLLM is running, if not start it and wait"""
    if check_vllm_connection():
        return None # Already running, no process to manage
    
    process = start_vllm_server()
    if not process:
        return None
        
    print("[System] Waiting for server startup (this may take a few minutes)...")
    # Wait loop with 5 min timeout
    max_retries = 60 # 60 * 5s = 300s
    for i in range(max_retries):
        if check_vllm_connection():
            print(f"\n[System] vLLM Server is ready!")
            return process
        
        # Check if process crashed
        if process.poll() is not None:
            print("\n[Error] vLLM process terminated unexpectedly.")
            return None
            
        time.sleep(5)
        print(".", end="", flush=True)
        
    print("\n[Error] Timeout waiting for vLLM server.")
    process.terminate()
    return None

def load_vendor_master():
    """Load vendor master data from Excel file"""
    path = os.path.join(SCRIPT_DIR, "Vendor_branch.xlsx")
    if not os.path.exists(path):
        return None
    try:
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        df['เลขประจำตัวผู้เสียภาษี'] = df['เลขประจำตัวผู้เสียภาษี'].fillna('').str.replace(r'\D', '', regex=True)
        df['สาขา'] = df['สาขา'].fillna('').str.strip()
        
        def clean_branch(x):
            x = str(x).strip()
            if x in ['สำนักงานใหญ่', 'สนญ', 'สนญ.', 'Head Office', 'H.O.', 'HO']:
                return '00000'
            if x.isdigit():
                return x.zfill(5)
            return x
        df['สาขา'] = df['สาขา'].apply(clean_branch)
        
        cols_to_return = ['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']
        if 'ชื่อบริษัท' in df.columns:
            cols_to_return.append('ชื่อบริษัท')
        return df[cols_to_return]
    except:
        return None

def get_target_pages(selection_str, total_pages):
    selection_str = str(selection_str).lower().replace(" ", "")
    if selection_str == 'all':
        return list(range(1, total_pages + 1))
    
    pages = set()
    for part in selection_str.split(','):
        if '-' in part:
            try:
                s, e = part.split('-')
                pages.update(range(int(s), (total_pages if e == 'n' else int(e)) + 1))
            except:
                pass
        elif part.isdigit():
            pages.add(int(part))
    return sorted([p for p in pages if 1 <= p <= total_pages])

def preprocess_image(image, max_size=1280):
    if image.mode != 'RGB':
        image = image.convert('RGB')
    image = ImageEnhance.Contrast(image).enhance(1.8)
    if max(image.size) > max_size:
        image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
    return image

# --- Updated Extract Function for vLLM (OpenAI Compatible) ---
def extract_text_from_image(file_path, pages_list):
    """Extract text from PDF pages using vLLM (Typhoon OCR)"""
    extracted_pages = []
    
    poppler = POPPLER_PATH if POPPLER_PATH and os.path.exists(POPPLER_PATH) else None
    
    is_pdf = file_path.lower().endswith('.pdf')
    if not is_pdf:
        pages_list = [1]

    for page_num in pages_list:
        try:
            print(f"   [Step 1] Rendering/Loading Page {page_num}...")
            
            if is_pdf:
                images = convert_from_path(
                    file_path,
                    first_page=page_num,
                    last_page=page_num,
                    poppler_path=poppler,
                    dpi=300
                )
                if not images:
                    continue
                img = preprocess_image(images[0])
            else:
                try:
                    img = Image.open(file_path)
                    img = preprocess_image(img)
                except Exception as img_err:
                    print(f"   [Error] Could not open image: {img_err}")
                    continue
            
            # Convert to Base64
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
            
            print(f"   [Step 2] Sending to AI (vLLM)...")
            
            # --- Payload for vLLM / OpenAI Compatible API ---
            payload = {
                "model": MODEL_NAME,
                "messages": [
                    # Explicit System Prompt to prevent hallucination
                    {
                        "role": "system", 
                        "content": "You are an OCR engine. Output only the text found in the image in Markdown format. Do not add any instructions, explanations, or conversational text."
                    },
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": "OCR this document. Extract all text, tables, and numbers exactly as shown. Output in Markdown. Do not include instructions or 'Here is the markdown'."},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{img_str}"
                                }
                            }
                        ]
                    }
                ],
                "max_tokens": 1024,
                "temperature": 0
            }
            
            headers = {"Content-Type": "application/json"}
            response = requests.post(VLLM_API_URL, headers=headers, json=payload, timeout=300)
            
            if response.status_code == 200:
                # --- Parse Response from OpenAI Format ---
                response_json = response.json()
                raw_content = response_json["choices"][0]["message"]["content"].strip()
                
                cleaned = clean_ocr_text(raw_content)
                extracted_pages.append((page_num, cleaned))
                print(f"   [Step 3] Page {page_num} Processed.")
            else:
                print(f"   [Error] API Status {response.status_code}: {response.text}")
            
            del img, img_str, buffered
            if is_pdf: del images
            gc.collect()
            
        except Exception as e:
            print(f"   [Error] Page {page_num}: {e}")
    
    return extracted_pages

def clean_ocr_text(text):
    if not text:
        return ""
    
    # List of common hallucination starts to strip
    garbage_headers = [
        "Formatting Rules:",
        "Only return the clean Markdown",
        "Here is the markdown",
        "Sure, here is the text",
        "Based on the image",
        "The text in the image is:",
        "Output:"
    ]
    
    # 1. Filter out known big blocks (like the Formatting Rules one)
    if "Formatting Rules:" in text and "Only return the clean Markdown" in text:
        # Remove everything up to "checked boxes." or end of instructions
        text = re.sub(r'^[\s\S]*?checked boxes\.\s*', '', text, flags=re.IGNORECASE)
        # Also try simpler cut if "checked boxes" isn't there
        text = re.sub(r'^[\s\S]*?Formatting Rules[\s\S]*?(\n\n|\r\n\r\n)', '', text, flags=re.IGNORECASE)

    # 2. Check for other conversational starters
    lines = text.split('\n')
    if lines:
        first_line = lines[0].strip()
        for header in garbage_headers:
            if header.lower() in first_line.lower():
                # If found, try to remove the first few lines until we hit data
                # Simple strategy: remove the first line
                lines = lines[1:]
                text = '\n'.join(lines)
                break
    
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n', '\n\n', text)
    return text.strip()

def main():
    print(f"--- OCR Processing (vLLM Local Mode) ---")
    print(f"Platform: {platform.system()} {platform.release()}")
    print(f"Source: {SOURCE_DIR}")
    print(f"Output: {OUTPUT_DIR}")
    print(f"API Endpoint: {VLLM_API_URL}")
    print(f"Model: {MODEL_NAME}")
    
    # --- Auto-Start Logic ---
    vllm_process = ensure_vllm_running()
    
    if not check_vllm_connection():
        print("[ERROR] Cannot connect to vLLM. Please ensure vLLM is running on port 8000.")
        if vllm_process: vllm_process.terminate()
        return
    
    templates = load_templates()
    if templates:
        available_types = list(templates.get("templates", {}).keys())
        print(f"Loaded templates: {available_types}")
    
    vendor_df = load_vendor_master()
    data_rows = []
    
    if not os.path.exists(SOURCE_DIR):
        print(f"[ERROR] Source directory not found: {SOURCE_DIR}")
        return
    
    files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg'))]
    
    if not files:
        print("No PDF files found.")
        return
    
    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        print(f"\n[File] {filename}")
        try:
            if filename.lower().endswith('.pdf'):
                reader = PdfReader(file_path)
                total_pages = len(reader.pages)
            else:
                total_pages = 1
                
            ocr_results = extract_text_from_image(
                file_path,
                get_target_pages(PAGE_CONFIG, total_pages)
            )
            
            for p_num, raw_text in ocr_results:
                txt_path = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_page{p_num}.txt")
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write(raw_text)
                
                parsed = parse_ocr_data_with_template(raw_text, templates, DOC_TYPE)
                print(f"   Detected Type: {parsed['document_type_name']}")
                
                row_data = {
                    "Link PDF": f'=HYPERLINK("{file_path}", "{filename}")',
                    "Page": p_num,
                    "Document Type": parsed["document_type_name"],
                    "VendorID_OCR": parsed["tax_id"],
                    "Branch_OCR": parsed["branch"],
                    "Document No": parsed["document_no"],
                    "Date": parsed["date"],
                    "Amount": parsed["amount"],
                }
                
                if parsed["document_type"] == "cy_instruction":
                    extra = parsed.get("extra_fields", {})
                    booking_no = extra.get("cy_booking", "")
                    row_data["CyOrg"] = extra.get("cy_org", "")
                    row_data["CyExporter"] = extra.get("cy_exporter", "")
                    row_data["CyInvoiceNo"] = extra.get("cy_invoice_no", "")
                    row_data["CyBooking"] = booking_no
                    cy_qty = extra.get("cy_qty", "")
                    row_data["CyQty"] = cy_qty
                    
                    containers = ""
                    if cy_qty:
                        qty_match = re.match(r'^([\d.]+)', cy_qty)
                        if qty_match:
                            try:
                                containers = str(int(float(qty_match.group(1))))
                            except:
                                pass
                    row_data["Containers"] = containers
                else:
                    for field_name, value in parsed.get("extra_fields", {}).items():
                        label = field_name.replace("_", " ").title()
                        row_data[label] = value
                
                data_rows.append(row_data)
                
        except Exception as e:
            print(f"   [Error] {filename}: {e}")

    if data_rows:
        df = pd.DataFrame(data_rows)
        if vendor_df is not None:
            print("\nMapping Vendor Code...")
            def clean_branch_code(val):
                s = str(val).strip()
                if s.lower() in ['nan', 'none', '']:
                    return "00000"
                if s.isdigit():
                    return s.zfill(5)
                return s

            if 'Branch_OCR' in df.columns:
                df['Branch_OCR'] = df['Branch_OCR'].apply(clean_branch_code)

            df = pd.merge(
                df,
                vendor_df,
                left_on=['VendorID_OCR', 'Branch_OCR'],
                right_on=['เลขประจำตัวผู้เสียภาษี', 'สาขา'],
                how='left'
            )
            df.rename(columns={'Vendor code SAP': 'Vendor code'}, inplace=True)
            df.drop(columns=['เลขประจำตัวผู้เสียภาษี', 'สาขา'], inplace=True, errors='ignore')
        else:
            df['Vendor code'] = ""
        
        priority_cols = [
            "Link PDF", "Page", "Document Type",
            "VendorID_OCR", "Branch_OCR", "Vendor code", "ชื่อบริษัท",
            "Document No", "Date", "Amount"
        ]
        
        all_cols = df.columns.tolist()
        final_cols = [col for col in priority_cols if col in all_cols]
        final_cols += [col for col in all_cols if col not in final_cols]
        
        df = df[final_cols]

        cy_columns = ['CyOrg', 'CyExporter', 'CyInvoiceNo', 'CyBooking', 'CyQty', 'Containers']
        for col in cy_columns:
            if col not in df.columns:
                df[col] = ""
        
        if 'Page' in df.columns:
            df = df.sort_values('Page', ascending=True).reset_index(drop=True)
        
        last_cy_values = {col: "" for col in cy_columns}
        
        for idx, row in df.iterrows():
            doc_type = str(row.get('Document Type', '')).strip().lower()
            if 'cy' in doc_type or 'instruction' in doc_type:
                for col in cy_columns:
                    if col in df.columns:
                        val = row.get(col, "")
                        if pd.notna(val) and str(val).strip():
                            last_cy_values[col] = val
            else:
                for col in cy_columns:
                    current_val = row.get(col, "")
                    if pd.isna(current_val) or str(current_val).strip() == "":
                        df.at[idx, col] = last_cy_values[col]

        output_excel_path = os.path.join(OUTPUT_DIR, "summary_ocr.xlsx")
        
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                sheet_name_mapping = {
                    'invoice': 'INVOICE',
                    'ใบกำกับภาษี/Invoice': 'INVOICE',
                    'Sahatthai Invoice': 'INVOICE',
                    'sahatthai_invoice': 'INVOICE',
                    'cy_instruction': 'CY_INSTRUCTION',
                    'CY INSTRUCTION': 'CY_INSTRUCTION',
                    'billing_note': 'ใบวางบิล',
                    'ใบวางบิล/Billing Note': 'ใบวางบิล'
                }
                
                if 'Document Type' in df.columns:
                    df['_sheet_name'] = df['Document Type'].apply(
                        lambda x: sheet_name_mapping.get(str(x).strip(), 'INVOICE') if pd.notna(x) else 'INVOICE'
                    )
                else:
                    df['_sheet_name'] = 'INVOICE'
                
                sheet_columns = {
                    'INVOICE': [
                        "Link PDF", "Page", "Document Type", 
                        "VendorID_OCR", "Branch_OCR", "Vendor code", "ชื่อบริษัท", 
                        "Document No", "Date", "Amount", 
                        "CyOrg", "CyExporter", "CyInvoiceNo", "CyBooking", "CyQty", "Containers"
                    ],
                    'CY_INSTRUCTION': [
                        "Link PDF", "Page", "Document Type", 
                        "CyOrg", "CyExporter", "CyInvoiceNo", "CyBooking", "CyQty", "Containers", "Container_delivery"
                    ],
                    'ใบวางบิล': [
                        "Link PDF", "Page", "Document Type", 
                        "VendorID_OCR", "Branch_OCR", "Vendor code", "ชื่อบริษัท", 
                        "Document No", "Date", "Amount"
                    ]
                }
                
                grouped = df.groupby('_sheet_name', dropna=False)
                
                invoice_df = grouped.get_group('INVOICE') if 'INVOICE' in [g[0] for g in grouped] else pd.DataFrame()
                cy_df_temp = grouped.get_group('CY_INSTRUCTION') if 'CY_INSTRUCTION' in [g[0] for g in grouped] else None
                
                if cy_df_temp is not None and not invoice_df.empty and 'CyInvoiceNo' in invoice_df.columns:
                    invoice_counts = invoice_df.groupby('CyInvoiceNo').size().to_dict()
                    def calc_container_delivery(cy_invoice_no):
                        if pd.isna(cy_invoice_no) or str(cy_invoice_no).strip() == '':
                            return ''
                        count = invoice_counts.get(str(cy_invoice_no).strip(), 0)
                        if count > 0:
                            return str(count * 0.5)
                        return ''
                    
                    cy_mask = df['_sheet_name'] == 'CY_INSTRUCTION'
                    if 'Container_delivery' not in df.columns:
                        df['Container_delivery'] = ''
                    df.loc[cy_mask, 'Container_delivery'] = df.loc[cy_mask, 'CyInvoiceNo'].apply(calc_container_delivery)
                    grouped = df.groupby('_sheet_name', dropna=False)
                
                for sheet_name, group_df in grouped:
                    if pd.isna(sheet_name) or str(sheet_name).strip() == '':
                        sheet_name = 'INVOICE'
                    
                    group_df = group_df.drop(columns=['_sheet_name'], errors='ignore')
                    target_cols = sheet_columns.get(str(sheet_name).strip(), sheet_columns['INVOICE'])
                    
                    available_cols = [col for col in target_cols if col in group_df.columns]
                    for col in target_cols:
                        if col not in available_cols:
                            group_df[col] = ""
                            available_cols.append(col)
                    
                    final_cols = [col for col in target_cols if col in group_df.columns]
                    group_df = group_df[final_cols]
                    group_df = group_df.reset_index(drop=True)
                    group_df.to_excel(writer, index=False, sheet_name=str(sheet_name))
                    print(f"   -> Sheet '{sheet_name}': {len(group_df)} rows")
                
            print(f"\nSuccess! Output saved at: {output_excel_path}")
            print(f"Total rows: {len(df)}")
        except Exception as e:
            print(f"Error saving Excel: {e}")
    else:
        print("No data extracted.")

if __name__ == "__main__":
    main()