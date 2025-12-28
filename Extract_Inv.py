import os
import sys
import requests
import json
import re
import pandas as pd
import platform
from pypdf import PdfReader

# --- Cross-platform Configuration ---
def get_default_source_dir():
    """Get default source directory based on OS"""
    if platform.system() == 'Windows':
        return r"d:\Project\ocr\source"
    else:
        return os.path.expanduser("~/ocr/source")


def get_default_output_dir():
    """Get default output directory based on OS"""
    if platform.system() == 'Windows':
        return r"d:\Project\ocr\output"
    else:
        return os.path.expanduser("~/ocr/output")


# --- Configuration (supports environment variables) ---
# API Key from environment variable (recommended) or fallback to config file
API_KEY = os.environ.get("TYPHOON_API_KEY", "")

# If not set in environment, try to load from config file
if not API_KEY:
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                API_KEY = config.get('API_KEY', '')
        except:
            pass

# Script directory for relative paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VENDOR_MASTER_FILE = "Vendor_branch.xlsx"

# Command line arguments or defaults
if len(sys.argv) >= 3:
    SOURCE_DIR = sys.argv[1]
    OUTPUT_DIR = sys.argv[2]
    PAGE_CONFIG = sys.argv[3] if len(sys.argv) > 3 else "All"
else:
    SOURCE_DIR = get_default_source_dir()
    OUTPUT_DIR = get_default_output_dir()
    PAGE_CONFIG = "2"

# Create output directory if not exists
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


# --- Load Vendor Master (Excel) ---
def load_vendor_master():
    """Load vendor master data from Excel file"""
    path = os.path.join(SCRIPT_DIR, VENDOR_MASTER_FILE)
    if not os.path.exists(path):
        print(f"Warning: Vendor master file not found: {VENDOR_MASTER_FILE} in {SCRIPT_DIR}")
        return None
    
    try:
        print(f"Loading Vendor Master from: {path}")
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        
        req_cols = ['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']
        if not all(col in df.columns for col in req_cols):
            print(f"Error: Missing columns in Master file (required: {req_cols})")
            return None

        df['เลขประจำตัวผู้เสียภาษี'] = df['เลขประจำตัวผู้เสียภาษี'].fillna('').str.replace(r'\D', '', regex=True)
        
        def clean_branch(x):
            x = str(x).strip()
            if x.isdigit():
                return x.zfill(5)
            return x
        
        df['สาขา'] = df['สาขา'].apply(clean_branch)
        
        return df[['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']]
        
    except Exception as e:
        print(f"Error reading Vendor file: {e}")
        return None


# --- Calculate target pages ---
def get_target_pages(selection_str, total_pages):
    """Parse page selection string and return list of pages to process"""
    selection_str = str(selection_str).lower().replace(" ", "")
    pages_to_process = set()

    if selection_str == 'all':
        return list(range(1, total_pages + 1))

    parts = selection_str.split(',')
    for part in parts:
        if '-' in part:
            start_s, end_s = part.split('-')
            start = int(start_s)
            end = total_pages if end_s == 'n' else int(end_s)
            end = min(end, total_pages)
            if start <= end:
                pages_to_process.update(range(start, end + 1))
        else:
            if part.isdigit():
                p = int(part)
                if 1 <= p <= total_pages:
                    pages_to_process.add(p)
    
    return sorted(list(pages_to_process))


# --- Typhoon OCR API ---
def extract_text_from_image(file_path, api_key, pages_list):
    """Extract text from PDF using Typhoon OCR API"""
    url = "https://api.opentyphoon.ai/v1/ocr"
    
    data = {
        'model': 'typhoon-ocr',
        'task_type': 'default', 
        'max_tokens': '16000',
        'temperature': '0.1',
        'top_p': '0.6',
        'repetition_penalty': '1.1'
    }
    
    if pages_list:
        data['pages'] = json.dumps(pages_list)
    
    headers = {'Authorization': f'Bearer {api_key}'}

    try:
        with open(file_path, 'rb') as file:
            files = {'file': file}
            response = requests.post(url, files=files, data=data, headers=headers)

        if response.status_code == 200:
            result = response.json()
            extracted_texts = []
            
            for page_result in result.get('results', []):
                if page_result.get('success'):
                    content = page_result['message']['choices'][0]['message']['content']
                    try:
                        parsed = json.loads(content)
                        text = parsed.get('natural_text', content)
                    except json.JSONDecodeError:
                        text = content
                    extracted_texts.append(text)
            
            return '\n'.join(extracted_texts)
        else:
            print(f"Error API: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Error processing file: {e}")
        return None


# --- Extract additional fields ---
def extract_additional_fields(text):
    """Extract additional fields from OCR text"""
    result = {
        'description': '',
        'sales_promotion': '',
        'total_amount': '',
        'withholding_tax': ''
    }
    
    if not text:
        return result
    
    # Extract description from HTML table
    description = ""
    
    if '<td>1</td><td>' in text:
        parts = text.split('<td>1</td><td>')
        if len(parts) > 1:
            after_item = parts[1]
            if '</td>' in after_item:
                description = after_item.split('</td>')[0]
    
    if not description:
        import re as regex_module
        match = regex_module.search(r'<td>(\d+)</td><td>(.+?)</td>', text, regex_module.DOTALL)
        if match:
            description = match.group(2)
    
    if description:
        description = re.sub(r'<br\s*/?>', ' ', description)
        description = re.sub(r'[\r\n]+', ' ', description)
        description = re.sub(r'<[^>]+>', '', description)
        description = re.sub(r'\s+', ' ', description).strip()
        
        if len(description) > 5:
            result['description'] = description
    
    # Sales promotion patterns
    sales_promo_patterns = [
        r'หมายเหตุ[^\n]*(?:\n|\r\n|\r)\s*(ค่าส่งเสริมการขาย[^\n]*)',
        r'ค่าส่งเสริมการขาย\s+(\d+%?\s*(?:จากยอด)?\s*[\d,]+\.?\d*)',
        r'ค่าส่งเสริมการขาย[^\d]*(?:\d+%?[^\n]*(?:จากยอด)?\s*[\d,]+\.?\d*)',
        r'ค่าส่งเสริมการขาย[^\n]+',
    ]
    
    for pattern in sales_promo_patterns:
        sales_match = re.search(pattern, text, re.IGNORECASE)
        if sales_match:
            sales_promo = sales_match.group(1) if sales_match.lastindex and sales_match.lastindex >= 1 else sales_match.group(0)
            sales_promo = re.sub(r'[\r\n]+', ' ', sales_promo)
            sales_promo = re.sub(r'\s+', ' ', sales_promo).strip()
            if sales_promo:
                result['sales_promotion'] = sales_promo
                break
    
    # Total amount patterns
    total_patterns = [
        r'รวมภาษีมูลค่าเพิ่ม\s*[:\.]?\s*([\d,]+\.\d{2})',
        r'รวม\s*ภาษีมูลค่าเพิ่ม[^\d]*([\d,]+\.\d{2})',
        r'รวมภาษีมูลค่าเพิ่ม[^\n]*(?:\n|\\n)\s*([\d,]+\.\d{2})',
    ]
    
    for pattern in total_patterns:
        total_match = re.search(pattern, text, re.IGNORECASE)
        if total_match:
            total_amount = total_match.group(1)
            if total_amount:
                result['total_amount'] = total_amount
                break
    
    # Withholding tax patterns
    withholding_patterns = [
        r'หัก\s+ณ\s+ที่จ่าย\s*[:\.]?\s*([\d,]+\.\d{2})',
        r'หัก\s*ณ\s*ที่จ่าย[^\d]*([\d,]+\.\d{2})',
        r'หัก\s*ณ\s*ที่จ่าย[^\n]*(?:\n|\\n)\s*([\d,]+\.\d{2})',
    ]
    
    for pattern in withholding_patterns:
        withholding_match = re.search(pattern, text, re.IGNORECASE)
        if withholding_match:
            withholding_tax = withholding_match.group(1)
            if withholding_tax:
                result['withholding_tax'] = withholding_tax
                break
    
    return result


# --- Parse OCR data ---
def parse_ocr_data(text):
    """Parse OCR text to extract structured data"""
    if not text:
        return None, None, None, None, None, None, None, None, None

    inv_match = re.search(r"เลขที่\s*[:\.]?\s*([A-Za-z0-9\-\/]{3,})", text)
    invoice_no = inv_match.group(1) if inv_match else ""

    date_match = re.search(r"วันที่\s*[:\.]?\s*(\d{1,2}\s+[^\s]+\s+\d{4}|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", text)
    bill_date = date_match.group(1) if date_match else ""

    amount_match = re.search(r"(?:จำนวนเงินรวมทั้งสิ้น|รวมเงินทั้งสิ้น|GRAND TOTAL)\s*[:\.]?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if amount_match:
        amount = amount_match.group(1)
    else:
        amounts = re.findall(r"([\d,]+\.\d{2})", text)
        amount = amounts[-1] if amounts else ""

    # Tax ID
    tax_id = ""
    all_tax_ids = re.findall(r"\b(\d{13})\b", text)
    
    if all_tax_ids:
        tax_id = all_tax_ids[0]
    else:
        tax_pattern_match = re.search(r"\b\d{1}-\d{4}-\d{5}-\d{2}-\d{1}\b", text)
        if tax_pattern_match:
            tax_id = re.sub(r"\D", "", tax_pattern_match.group(0))
        else:
            tax_keyword_match = re.search(r"(?:เลข(?:ที่)?(?:ประจำตัว)?ผู้เสียภาษี(?:อากร)?|Tax\s*ID|Tax\s*No\.?|เลขทะเบียนนิติบุคคล)\s*[:\.]?\s*([0-9\-\s]{10,25})", text, re.IGNORECASE)
            if tax_keyword_match:
                clean_tax = re.sub(r"\D", "", tax_keyword_match.group(1))
                if len(clean_tax) >= 10:
                    tax_id = clean_tax[:13] if len(clean_tax) >= 13 else clean_tax

    # Branch ID
    branch_id = ""
    ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
    if ho_match:
        branch_id = "00000"
    else:
        branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})", text, re.IGNORECASE)
        if branch_match:
            branch_id = branch_match.group(1).zfill(5)
    
    # Additional fields
    additional_fields = extract_additional_fields(text)
    description = additional_fields['description']
    sales_promotion = additional_fields['sales_promotion']
    total_amount = additional_fields['total_amount']
    withholding_tax = additional_fields['withholding_tax']

    return invoice_no, bill_date, amount, tax_id, branch_id, description, sales_promotion, total_amount, withholding_tax


# --- Main Logic ---
def main():
    print(f"--- Start Processing (Mode: {PAGE_CONFIG}) ---")
    print(f"Platform: {platform.system()} {platform.release()}")
    print(f"Source: {SOURCE_DIR}")
    print(f"Output: {OUTPUT_DIR}")
    
    if not API_KEY:
        print("[ERROR] API Key not set. Please set TYPHOON_API_KEY environment variable or update config.json")
        return
    
    # Load Vendor Master
    vendor_df = load_vendor_master()
    
    data_rows = []
    
    if not os.path.exists(SOURCE_DIR):
        print(f"Error: Source directory not found: {SOURCE_DIR}")
        return

    files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".pdf")]
    
    if not files:
        print("No PDF files found.")
        return

    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        print(f"Processing: {filename}")

        try:
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)
            target_pages = get_target_pages(PAGE_CONFIG, total_pages)
            print(f"   -> Total Pages: {total_pages}, Target: {target_pages}")

            for page_num in target_pages:
                print(f"      Reading Page {page_num}...")
                page_text = extract_text_from_image(file_path, API_KEY, pages_list=[page_num])
                
                if page_text:
                    txt_filename = f"{os.path.splitext(filename)[0]}_page{page_num}.txt"
                    txt_path = os.path.join(OUTPUT_DIR, txt_filename)
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write(page_text)
                    
                    invoice, date_val, money, tax_id_val, branch_val, description, sales_promotion, total_amount, withholding_tax = parse_ocr_data(page_text)
                    hyperlink_formula = f'=HYPERLINK("{file_path}", "{filename} (Page {page_num})")'

                    data_rows.append({
                        "Link PDF": hyperlink_formula,
                        "Page": page_num,
                        "VendorID_OCR": tax_id_val,
                        "Branch_OCR": branch_val,
                        "Invoice No": invoice,
                        "Date": date_val,
                        "Amount": money,
                        "Description": description,
                        "Sales Promotion": sales_promotion,
                        "Total Amount": total_amount,
                        "Withholding Tax": withholding_tax
                    })
                else:
                    print(f"      Warning: Failed to read page {page_num}")

        except Exception as e:
            print(f"   Error reading PDF file: {e}")

    # Save and merge data
    if data_rows:
        df = pd.DataFrame(data_rows)
        
        if vendor_df is not None:
            print("Mapping Vendor Code...")
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

        columns_order = [
            "Link PDF", "Page", 
            "VendorID_OCR", "Branch_OCR", "Vendor code", 
            "Invoice No", "Date", "Amount",
            "Description", "Sales Promotion", "Total Amount", "Withholding Tax"
        ]
        
        final_cols = [col for col in columns_order if col in df.columns]
        df = df[final_cols]

        output_excel_path = os.path.join(OUTPUT_DIR, "summary_ocr.xlsx")
        
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            print(f"Success! Output saved at: {output_excel_path}")
        except Exception as e:
            print(f"Error saving Excel: {e}")
    else:
        print("No data extracted.")


if __name__ == "__main__":
    main()
