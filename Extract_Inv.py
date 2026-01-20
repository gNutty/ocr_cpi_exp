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
TEMPLATES_FILE = "document_templates.json"

# Command line arguments or defaults
# Usage: python Extract_Inv.py <source_dir> <output_dir> <page_config> [document_type]
if len(sys.argv) >= 3:
    SOURCE_DIR = sys.argv[1]
    OUTPUT_DIR = sys.argv[2]
    PAGE_CONFIG = sys.argv[3] if len(sys.argv) > 3 else "All"
    DOC_TYPE = sys.argv[4] if len(sys.argv) > 4 else "auto"
else:
    SOURCE_DIR = get_default_source_dir()
    OUTPUT_DIR = get_default_output_dir()
    PAGE_CONFIG = "2"
    DOC_TYPE = "auto"

# Create output directory if not exists
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


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
    """Auto-detect document type based on keywords in text.
    Only 3 document types are supported:
    1. billing_note (ใบวางบิล)
    2. cy_instruction (CY INSTRUCTION)
    3. invoice (ใบกำกับภาษี/Invoice) - default for everything else
    """
    if not text or not templates:
        return "invoice"  # default
    
    text_lower = text.lower()
    
    # Check first 15 lines (header) for document title keywords
    lines = text.split('\n')
    header_text = "\n".join(lines[:15]).lower()
    
    # Strict Header: First 4 lines (Higher priority)
    strict_header = "\n".join(lines[:4]).lower()
    
    # Priority 1: Check for CY INSTRUCTION (unique keywords)
    cy_template = templates.get("templates", {}).get("cy_instruction", {})
    cy_keywords = cy_template.get("detect_keywords", [])
    for keyword in cy_keywords:
        if keyword.lower() in text_lower:
            return "cy_instruction"
    
    # Priority 2: Check for ใบวางบิล (Billing Note)
    billing_template = templates.get("templates", {}).get("billing_note", {})
    billing_keywords = billing_template.get("detect_keywords", [])
    for keyword in billing_keywords:
        kw_low = keyword.lower()
        # Check in strict header first (top 4 lines)
        if kw_low in strict_header:
            return "billing_note"
        # Then check in general header
        if kw_low in header_text:
            return "billing_note"
    
    # Priority 3: Check for Sahatthai Invoice (special case, maps to invoice)
    sahatthai_template = templates.get("templates", {}).get("sahatthai_invoice", {})
    sahatthai_keywords = sahatthai_template.get("detect_keywords", [])
    for keyword in sahatthai_keywords:
        if keyword.lower() in text_lower:
            return "sahatthai_invoice"  # Will be displayed as Invoice
    
    # Default: Everything else is Invoice
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
                # Get first capturing group or full match
                value = match.group(1) if match.lastindex and match.lastindex >= 1 else match.group(0)
                
                # Clean HTML if specified
                if options.get("clean_html"):
                    value = re.sub(r'<br\s*/?>', ' ', value)
                    value = re.sub(r'<[^>]+>', '', value)
                
                # Clean whitespace
                value = re.sub(r'[\r\n]+', ' ', value)
                value = re.sub(r'\s+', ' ', value).strip()
                
                # Smart extraction for booking numbers: find pattern like "E BKG13808784"
                # Extract only text that starts with letters and ends with digits
                if options.get("extract_booking_pattern"):
                    # Split into words and collect: letter-only words followed by alphanumeric ending in digit
                    words = value.split()
                    result_parts = []
                    found_number = False
                    
                    for word in words:
                        # Pure letters (like 'E' or 'BKG')
                        if re.match(r'^[A-Za-z]+$', word):
                            if not found_number:
                                result_parts.append(word)
                        # Alphanumeric ending with digits (like 'BKG13808784')
                        elif re.match(r'^[A-Za-z]+\d+$', word):
                            result_parts.append(word)
                            found_number = True
                            break  # Stop after finding the number part
                        # Pure digits
                        elif re.match(r'^\d+$', word):
                            result_parts.append(word)
                            found_number = True
                            break
                        else:
                            # If we already started collecting and hit something else, stop
                            if result_parts:
                                break
                    
                    if result_parts:
                        value = ' '.join(result_parts)
                
                # specific option to remove all spaces (e.g. for Booking No)
                if options.get("remove_spaces"):
                    value = value.replace(" ", "")
                
                # Check minimum digits requirement (for document_no validation)
                if options.get("min_digits"):
                    digit_count = len(re.sub(r'\D', '', value))
                    if digit_count < options["min_digits"]:
                        continue  # Skip to next pattern
                
                # Clean non-digits if specified
                if options.get("clean_non_digits"):
                    value = re.sub(r'\D', '', value)
                    # Truncate to specified length
                    if options.get("length"):
                        value = value[:options["length"]]
                
                if value:
                    return value
        except Exception:
            continue
    
    return ""


def extract_common_fields(text, common_fields_config, doc_type_name=None):
    """Extract common fields (tax_id, branch) that apply to all document types"""
    result = {"tax_id": "", "branch": ""}
    
    if not text or not common_fields_config:
        return result
    
    # Extract Tax ID
    tax_config = common_fields_config.get("tax_id", {})
    tax_patterns = tax_config.get("patterns", [])
    
    # Company Tax ID to skip (always extract vendor's Tax ID, not company's)
    COMPANY_TAX_ID = "0105522018355"
    
    # Method 1: Find all 13-digit numbers directly
    all_tax_ids = re.findall(r"\b(\d{13})\b", text)
    vendor_tax_ids = [tid for tid in all_tax_ids if tid != COMPANY_TAX_ID]
    
    if vendor_tax_ids:
        result["tax_id"] = vendor_tax_ids[0]
    else:
        # Method 2: Try pattern with dashes (e.g., 0-1234-56789-01-2)
        all_dashed = re.findall(r"\b(\d{1}-\d{4}-\d{5}-\d{2}-\d{1})\b", text)
        for match in all_dashed:
            clean_id = re.sub(r"\D", "", match)
            if clean_id != COMPANY_TAX_ID:
                result["tax_id"] = clean_id
                break
        
        # Method 3: Try pattern with spaces (e.g., 0 123456789012)
        if not result["tax_id"]:
            spaced_matches = re.findall(r"\b(\d{1}\s+\d{12})\b", text)
            for match in spaced_matches:
                clean_id = re.sub(r"\D", "", match)
                if clean_id != COMPANY_TAX_ID:
                    result["tax_id"] = clean_id
                    break
        
        # Method 4: Keyword-based extraction
        if not result["tax_id"]:
            for pattern in tax_patterns:
                value = extract_field_by_patterns(text, [pattern], {"clean_non_digits": True, "length": 13})
                if value and len(value) >= 10 and value != COMPANY_TAX_ID:
                    result["tax_id"] = value
                    break
    
    # Extract Branch
    branch_config = common_fields_config.get("branch", {})
    branch_patterns = branch_config.get("patterns", [])
    default_hq = branch_config.get("default_hq", "00000")
    pad_zeros = branch_config.get("pad_zeros", 5)
    
    # Try to find specific branch number FIRST (prioritize over Head Office)
    branch_match = re.search(r"(?:สาขา(?:ที่)?(?:ออกใบกำกับภาษี)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})(?!\d)", text, re.IGNORECASE)
    if branch_match:
        result["branch"] = branch_match.group(1).zfill(pad_zeros)
    else:
        # Fall back to Head Office keywords
        ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
        if ho_match:
            result["branch"] = default_hq
    
    return result


def parse_ocr_data_with_template(text, templates, doc_type="auto"):
    """Parse OCR text using document template patterns"""
    result = {
        "document_type": "",
        "document_type_name": "",
        "document_no": "",
        "date": "",
        "amount": "",
        "tax_id": "",
        "branch": "",
        "extra_fields": {}
    }
    
    if not text:
        return result
    
    # Load templates if not provided
    if not templates:
        templates = load_templates()
    
    if not templates:
        # Fallback to basic extraction
        return parse_ocr_data_basic(text)
    
    # Detect or use specified document type
    if doc_type == "auto":
        detected_type = detect_document_type(text, templates)
    else:
        detected_type = doc_type if doc_type in templates.get("templates", {}) else "invoice"
    
    result["document_type"] = detected_type
    
    # Get template for this document type
    template = templates.get("templates", {}).get(detected_type, {})
    result["document_type_name"] = template.get("name", detected_type)
    
    # Extract common fields (tax_id, branch) - always extracted for Vendor lookup
    common_fields = templates.get("common_fields", {})
    common_result = extract_common_fields(text, common_fields, result["document_type_name"])
    result["tax_id"] = common_result["tax_id"]
    result["branch"] = common_result["branch"]
    
    # Extract template-specific fields
    fields_config = template.get("fields", {})
    
    for field_name, field_config in fields_config.items():
        patterns = field_config.get("patterns", [])
        options = {
            "clean_html": field_config.get("clean_html", False),
            "clean_non_digits": field_config.get("clean_non_digits", False),
            "length": field_config.get("length"),
            "min_digits": field_config.get("min_digits"),
            "remove_spaces": field_config.get("remove_spaces", False),
            "extract_booking_pattern": field_config.get("extract_booking_pattern", False)
        }
        
        # Handle skip_lines: skip first N lines before searching
        text_to_search = text
        skip_lines = field_config.get("skip_lines", 0)
        if skip_lines > 0:
            lines = text.split('\n')
            text_to_search = '\n'.join(lines[skip_lines:])
        
        value = extract_field_by_patterns(text_to_search, patterns, options)
        
        # Handle fallback for amount fields
        if not value and field_config.get("fallback") == "last_amount":
            amounts = re.findall(r"([\d,]+\.\d{2})", text)
            value = amounts[-1] if amounts else ""
        
        # Reject 13-digit numbers for document_no (these are Tax IDs, not document numbers)
        if field_name == "document_no" and value:
            # Remove non-alphanumeric for checking
            digits_only = re.sub(r'\D', '', value)
            if len(digits_only) == 13 and digits_only.isdigit():
                # This is likely a Tax ID, not a document number
                value = ""
        
        # Store in appropriate location
        if field_name in ["document_no", "date", "amount"]:
            result[field_name] = value
        else:
            result["extra_fields"][field_name] = value
    
    return result


def parse_ocr_data_basic(text):
    """Basic OCR parsing without templates (fallback)"""
    result = {
        "document_type": "invoice",
        "document_type_name": "ใบกำกับภาษี/Invoice",
        "document_no": "",
        "date": "",
        "amount": "",
        "tax_id": "",
        "branch": "",
        "extra_fields": {}
    }
    
    if not text:
        return result
    
    # Document number
    inv_match = re.search(r"เลขที่\s*[:\.]?\s*([A-Za-z0-9\-\/]{3,})", text)
    result["document_no"] = inv_match.group(1) if inv_match else ""
    
    # Date
    date_match = re.search(r"วันที่\s*[:\.]?\s*(\d{1,2}\s+[^\s]+\s+\d{4}|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", text)
    result["date"] = date_match.group(1) if date_match else ""
    
    # Amount
    amount_match = re.search(r"(?:จำนวนเงินรวมทั้งสิ้น|รวมเงินทั้งสิ้น|GRAND TOTAL)\s*[:\.]?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if amount_match:
        result["amount"] = amount_match.group(1)
    else:
        amounts = re.findall(r"([\d,]+\.\d{2})", text)
        result["amount"] = amounts[-1] if amounts else ""
    
    # Tax ID
    all_tax_ids = re.findall(r"\b(\d{13})\b", text)
    if all_tax_ids:
        result["tax_id"] = all_tax_ids[0]
    else:
        tax_pattern_match = re.search(r"\b\d{1}-\d{4}-\d{5}-\d{2}-\d{1}\b", text)
        if tax_pattern_match:
            result["tax_id"] = re.sub(r"\D", "", tax_pattern_match.group(0))
    
    # Branch
    ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
    if ho_match:
        result["branch"] = "00000"
    else:
        branch_match = re.search(r"(?:สาขา(?:ที่)?(?:ออกใบกำกับภาษี)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})(?!\d)", text, re.IGNORECASE)
        if branch_match:
            result["branch"] = branch_match.group(1).zfill(5)
    
    return result


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
            # Convert head office keywords to 00000
            if x in ['สำนักงานใหญ่', 'สนญ', 'สนญ.', 'Head Office', 'H.O.', 'HO']:
                return '00000'
            if x.isdigit():
                return x.zfill(5)
            return x
        
        df['สาขา'] = df['สาขา'].apply(clean_branch)
        
        # Also get company name if available
        cols_to_return = ['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']
        if 'ชื่อบริษัท' in df.columns:
            cols_to_return.append('ชื่อบริษัท')
        
        return df[cols_to_return]
        
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
    
    is_pdf = file_path.lower().endswith('.pdf')
    if pages_list and is_pdf:
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
                    
                    # Clean tags immediately
                    text = clean_ocr_text(text)
                    extracted_texts.append(text)
            
            return '\n'.join(extracted_texts)
        else:
            print(f"Error API: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Error processing file: {e}")
        return None


def clean_ocr_text(text):
    """Clean OCR extracted text"""
    if not text:
        return ""
    # Remove HTML tags including <?> and <table>
    text = re.sub(r'<[^>]+>', ' ', text)
    # Normalize whitespace
    text = re.sub(r'[\r\n]+', '\n', text)
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()


# --- Main Logic ---
def main():
    print(f"--- Start Processing ---")
    print(f"Platform: {platform.system()} {platform.release()}")
    print(f"Source: {SOURCE_DIR}")
    print(f"Output: {OUTPUT_DIR}")
    print(f"Page Config: {PAGE_CONFIG}")
    print(f"Document Type: {DOC_TYPE}")
    
    if not API_KEY:
        print("[ERROR] API Key not set. Please set TYPHOON_API_KEY environment variable or update config.json")
        return
    
    # Load templates
    templates = load_templates()
    if templates:
        available_types = list(templates.get("templates", {}).keys())
        print(f"Loaded templates: {available_types}")
    
    # Load Vendor Master
    vendor_df = load_vendor_master()
    
    data_rows = []
    
    if not os.path.exists(SOURCE_DIR):
        print(f"Error: Source directory not found: {SOURCE_DIR}")
        return

    files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg'))]
    
    if not files:
        print("No PDF files found.")
        return

    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        print(f"\nProcessing: {filename}")

        try:
            if filename.lower().endswith('.pdf'):
                reader = PdfReader(file_path)
                total_pages = len(reader.pages)
            else:
                total_pages = 1
                
            target_pages = get_target_pages(PAGE_CONFIG, total_pages)
            print(f"   -> Total Pages: {total_pages}, Target: {target_pages}")

            for page_num in target_pages:
                print(f"      Reading Page {page_num}...")
                page_text = extract_text_from_image(file_path, API_KEY, pages_list=[page_num])
                
                if page_text:
                    # Save raw OCR text
                    txt_filename = f"{os.path.splitext(filename)[0]}_page{page_num}.txt"
                    txt_path = os.path.join(OUTPUT_DIR, txt_filename)
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write(page_text)
                    
                    # Parse using templates
                    parsed = parse_ocr_data_with_template(page_text, templates, DOC_TYPE)
                    
                    print(f"      Detected Type: {parsed['document_type_name']}")
                    
                    hyperlink_formula = f'=HYPERLINK("{file_path}", "{filename} (Page {page_num})")'

                    row_data = {
                        "Link PDF": hyperlink_formula,
                        "Page": page_num,
                        "Document Type": parsed["document_type_name"],
                        "VendorID_OCR": parsed["tax_id"],
                        "Branch_OCR": parsed["branch"],
                        "Document No": parsed["document_no"],
                        "Date": parsed["date"],
                        "Amount": parsed["amount"],
                    }
                    
                    # Special handling for CY INSTRUCTION document type
                    if parsed["document_type"] == "cy_instruction":
                        extra = parsed.get("extra_fields", {})
                        
                        # Build CyBooking field: Only booking number, no extra fields
                        booking_no = extra.get("cy_booking", "")
                        cy_booking = booking_no
                        
                        # Add CY-specific columns
                        row_data["CyOrg"] = extra.get("cy_org", "")
                        row_data["CyExporter"] = extra.get("cy_exporter", "")
                        row_data["CyInvoiceNo"] = extra.get("cy_invoice_no", "")
                        row_data["CyBooking"] = cy_booking
                        cy_qty = extra.get("cy_qty", "")
                        row_data["CyQty"] = cy_qty
                        
                        # Extract container count from CyQty
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
                        # Add extra fields from template for other document types
                        for field_name, value in parsed.get("extra_fields", {}).items():
                            # Convert field_name to readable label
                            label = field_name.replace("_", " ").title()
                            row_data[label] = value
                    
                    data_rows.append(row_data)
                else:
                    print(f"      Warning: Failed to read page {page_num}")

        except Exception as e:
            print(f"   Error reading PDF file: {e}")

    # Save and merge data
    if data_rows:
        df = pd.DataFrame(data_rows)
        
        # Remap "Sahatthai Invoice" to "ใบกำกับภาษี/Invoice" for Summary Excel
        if 'Document Type' in df.columns:
            df['Document Type'] = df['Document Type'].replace({
                'Sahatthai Invoice': 'ใบกำกับภาษี/Invoice',
                'sahatthai_invoice': 'ใบกำกับภาษี/Invoice'
            })
        
        if vendor_df is not None:
            print("\nMapping Vendor Code...")
            # Helper to clean branch
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
            df.rename(columns={'Vendor code SAP': 'Vendor code', 'ชื่อบริษัท': 'Vendor Name'}, inplace=True)
            df.drop(columns=['เลขประจำตัวผู้เสียภาษี', 'สาขา'], inplace=True, errors='ignore')
        else:
            df['Vendor code'] = ""

        # Reorder columns - put important ones first
        priority_cols = [
            "Link PDF", "Page", "Document Type",
            "VendorID_OCR", "Branch_OCR", "Vendor code", "Vendor Name",
            "Document No", "Date", "Amount"
        ]
        
        # Get all columns, prioritizing the defined order
        all_cols = df.columns.tolist()
        final_cols = [col for col in priority_cols if col in all_cols]
        final_cols += [col for col in all_cols if col not in final_cols]
        
        df = df[final_cols]

        # ============================================================
        # CY Lookup: Forward-fill CY columns from CY INSTRUCTION to INVOICE
        # For each Invoice row, copy CY values from the most recent 
        # CY INSTRUCTION document that appeared before it (by page order)
        # ============================================================
        cy_columns = ['CyOrg', 'CyExporter', 'CyInvoiceNo', 'CyBooking', 'CyQty', 'Containers']
        
        # Ensure CY columns exist
        for col in cy_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Sort by Page to ensure correct order
        if 'Page' in df.columns:
            df = df.sort_values('Page', ascending=True).reset_index(drop=True)
        
        # Track last CY INSTRUCTION values
        last_cy_values = {col: "" for col in cy_columns}
        
        # Iterate through rows and apply lookup
        for idx, row in df.iterrows():
            doc_type = str(row.get('Document Type', '')).strip().lower()
            
            # Check if this is a CY INSTRUCTION
            if 'cy' in doc_type or 'instruction' in doc_type:
                # Update tracked values from this CY INSTRUCTION
                for col in cy_columns:
                    if col in df.columns:
                        val = row.get(col, "")
                        if pd.notna(val) and str(val).strip():
                            last_cy_values[col] = val
            else:
                # For other documents (Invoice, etc.), copy the tracked CY values
                for col in cy_columns:
                    current_val = row.get(col, "")
                    # Only fill if the current value is empty
                    if pd.isna(current_val) or str(current_val).strip() == "":
                        df.at[idx, col] = last_cy_values[col]
        
        print(f"Applied CY lookup to {len(df)} rows")

        output_excel_path = os.path.join(OUTPUT_DIR, "summary_ocr.xlsx")
        
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                # Define sheet name mapping for document types
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
                
                # Map document types to sheet names
                if 'Document Type' in df.columns:
                    df['_sheet_name'] = df['Document Type'].apply(
                        lambda x: sheet_name_mapping.get(str(x).strip(), 'INVOICE') if pd.notna(x) else 'INVOICE'
                    )
                else:
                    df['_sheet_name'] = 'INVOICE'
                
                # Define column configuration for each sheet
                sheet_columns = {
                    'INVOICE': [
                        "Link PDF", "Page", "Document Type", 
                        "VendorID_OCR", "Branch_OCR", "Vendor code", "Vendor Name", 
                        "Document No", "Date", "Amount", 
                        "CyOrg", "CyExporter", "CyInvoiceNo", "CyBooking", "CyQty", "Containers"
                    ],
                    'CY_INSTRUCTION': [
                        "Link PDF", "Page", "Document Type", 
                        "CyOrg", "CyExporter", "CyInvoiceNo", "CyBooking", "CyQty", "Containers", "Container_delivery"
                    ],
                    'ใบวางบิล': [
                        "Link PDF", "Page", "Document Type", 
                        "VendorID_OCR", "Branch_OCR", "Vendor code", "Vendor Name", 
                        "Document No", "Date", "Amount"
                    ]
                }
                
                # Group by sheet name and save each group to a separate sheet
                grouped = df.groupby('_sheet_name', dropna=False)
                
                # ============================================================
                # Calculate Container_delivery for CY_INSTRUCTION:
                # Count INVOICE rows with matching CyInvoiceNo * 0.5
                # ============================================================
                invoice_df = grouped.get_group('INVOICE') if 'INVOICE' in [g[0] for g in grouped] else pd.DataFrame()
                cy_df_temp = grouped.get_group('CY_INSTRUCTION') if 'CY_INSTRUCTION' in [g[0] for g in grouped] else None
                
                if cy_df_temp is not None and not invoice_df.empty and 'CyInvoiceNo' in invoice_df.columns:
                    # Count invoices per CyInvoiceNo
                    invoice_counts = invoice_df.groupby('CyInvoiceNo').size().to_dict()
                    
                    # Calculate Container_delivery for each CY row
                    def calc_container_delivery(cy_invoice_no):
                        if pd.isna(cy_invoice_no) or str(cy_invoice_no).strip() == '':
                            return ''
                        count = invoice_counts.get(str(cy_invoice_no).strip(), 0)
                        if count > 0:
                            return str(count * 0.5)
                        return ''
                    
                    # Apply calculation to original df for CY_INSTRUCTION rows
                    cy_mask = df['_sheet_name'] == 'CY_INSTRUCTION'
                    if 'Container_delivery' not in df.columns:
                        df['Container_delivery'] = ''
                    df.loc[cy_mask, 'Container_delivery'] = df.loc[cy_mask, 'CyInvoiceNo'].apply(calc_container_delivery)
                    
                    # Re-group after adding Container_delivery
                    grouped = df.groupby('_sheet_name', dropna=False)
                
                for sheet_name, group_df in grouped:
                    if pd.isna(sheet_name) or str(sheet_name).strip() == '':
                        sheet_name = 'INVOICE'
                    
                    # Drop the temporary _sheet_name column
                    group_df = group_df.drop(columns=['_sheet_name'], errors='ignore')
                    
                    # Select specific columns for this sheet
                    target_cols = sheet_columns.get(str(sheet_name).strip(), sheet_columns['INVOICE'])
                    
                    # Filter existing columns only
                    available_cols = [col for col in target_cols if col in group_df.columns]
                    
                    # Add missing columns as empty (optional, but good for consistency)
                    for col in target_cols:
                        if col not in available_cols:
                            group_df[col] = ""
                            available_cols.append(col)
                    
                    # Reorder columns matches the target list
                    # Use a set for faster lookup to avoid duplicates if any
                    final_cols = [col for col in target_cols if col in group_df.columns]
                    
                    # Final filtering and reordering
                    group_df = group_df[final_cols]
                    
                    # Reset index for clean output
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
