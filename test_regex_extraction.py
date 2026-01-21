"""
Test Regex Extraction Script
=============================
This script reads all .txt files from the example_doc folder,
applies regex patterns from document_templates.json, and
outputs the results to summary_ocr.xlsx for verification.

This script matches the output format and calculations of Extract_Inv.py.

Usage: python test_regex_extraction.py
"""

import os
import re
import json
import pandas as pd
from datetime import datetime

# Script directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXAMPLE_DOC_DIR = os.path.join(SCRIPT_DIR, "example_doc")
TEMPLATES_FILE = os.path.join(SCRIPT_DIR, "document_templates.json")
VENDOR_FILE = os.path.join(SCRIPT_DIR, "Vendor_branch.xlsx")
OUTPUT_FILE = os.path.join(EXAMPLE_DOC_DIR, "summary_ocr.xlsx")


def load_templates():
    """Load document templates from JSON file"""
    if not os.path.exists(TEMPLATES_FILE):
        print(f"Error: Templates file not found: {TEMPLATES_FILE}")
        return None
    
    with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_vendor_master():
    """Load vendor master data from Excel file"""
    if not os.path.exists(VENDOR_FILE):
        print(f"Warning: Vendor file not found: {VENDOR_FILE}")
        return None
    
    try:
        df = pd.read_excel(VENDOR_FILE, dtype=str)
        df.columns = df.columns.str.strip()
        df['เลขประจำตัวผู้เสียภาษี'] = df['เลขประจำตัวผู้เสียภาษี'].fillna('').str.replace(r'\D', '', regex=True)
        
        def clean_branch(x):
            x = str(x).strip()
            if x in ['สำนักงานใหญ่', 'สนญ', 'สนญ.', 'Head Office', 'H.O.', 'HO']:
                return '00000'
            if x.isdigit():
                return x.zfill(5)
            return x
        
        df['สาขา'] = df['สาขา'].apply(clean_branch)
        
        cols = ['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']
        if 'ชื่อบริษัท' in df.columns:
            cols.append('ชื่อบริษัท')
        
        return df[cols]
    except Exception as e:
        print(f"Error loading vendor file: {e}")
        return None


def detect_document_type(text, templates):
    """Auto-detect document type based on keywords in text"""
    if not text or not templates:
        return "invoice"
    
    text_lower = text.lower()
    lines = text.split('\n')
    header_text = "\n".join(lines[:15]).lower()
    strict_header = "\n".join(lines[:4]).lower()
    
    # Priority 1: CY INSTRUCTION
    cy_template = templates.get("templates", {}).get("cy_instruction", {})
    for keyword in cy_template.get("detect_keywords", []):
        if keyword.lower() in text_lower:
            return "cy_instruction"
    
    # Priority 2: Billing Note
    billing_template = templates.get("templates", {}).get("billing_note", {})
    for keyword in billing_template.get("detect_keywords", []):
        kw_low = keyword.lower()
        if kw_low in strict_header or kw_low in header_text:
            return "billing_note"
    
    # Priority 3: Sahatthai Invoice
    sahatthai_template = templates.get("templates", {}).get("sahatthai_invoice", {})
    for keyword in sahatthai_template.get("detect_keywords", []):
        if keyword.lower() in text_lower:
            return "sahatthai_invoice"
    
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
                
                # Clean HTML
                if options.get("clean_html"):
                    value = re.sub(r'<br\s*/?>', ' ', value)
                    value = re.sub(r'<[^>]+>', '', value)
                
                # Clean whitespace
                value = re.sub(r'[\r\n]+', ' ', value)
                value = re.sub(r'\s+', ' ', value).strip()
                
                # Extract booking pattern
                if options.get("extract_booking_pattern"):
                    words = value.split()
                    result_parts = []
                    found_number = False
                    
                    for word in words:
                        if re.match(r'^[A-Za-z]+$', word):
                            if not found_number:
                                result_parts.append(word)
                        elif re.match(r'^[A-Za-z]+\d+$', word):
                            result_parts.append(word)
                            found_number = True
                            break
                        elif re.match(r'^\d+$', word):
                            result_parts.append(word)
                            found_number = True
                            break
                        else:
                            if result_parts:
                                break
                    
                    if result_parts:
                        value = ' '.join(result_parts)
                
                # Remove spaces
                if options.get("remove_spaces"):
                    value = value.replace(" ", "")
                
                # Check minimum digits
                if options.get("min_digits"):
                    digit_count = len(re.sub(r'\D', '', value))
                    if digit_count < options["min_digits"]:
                        continue
                
                # Clean non-digits
                if options.get("clean_non_digits"):
                    value = re.sub(r'\D', '', value)
                    if options.get("length"):
                        value = value[:options["length"]]
                
                if value:
                    return value
        except Exception:
            continue
    
    return ""


def extract_common_fields(text, common_fields_config):
    """Extract common fields (tax_id, branch)"""
    result = {"tax_id": "", "branch": ""}
    
    if not text or not common_fields_config:
        return result
    
    COMPANY_TAX_ID = "0105522018355"
    
    # Special condition: if document contains "สยามคอนเทนเนอร์ เทอร์มินอล", use their Tax ID
    SIAM_CONTAINER_TAX_ID = "0105531101901"
    # Special condition: if document contains "สหไทย เทอร์มินอล", use their Tax ID
    SAHATHAI_TAX_ID = "0107560000192"
    # Special condition: if document contains "มนต์โลจิสติกส์ เซอร์วิส", use their Tax ID
    MON_LOGISTICS_TAX_ID = "0105559135291"
    
    if "สยามคอนเทนเนอร์ เทอร์มินอล" in text or "สยามคอนเทนเนอร์เทอร์มินอล" in text:
        result["tax_id"] = SIAM_CONTAINER_TAX_ID
    elif "สหไทย เทอร์มินอล" in text or "สหไทยเทอร์มินอล" in text:
        result["tax_id"] = SAHATHAI_TAX_ID
    elif "มนต์โลจิสติกส์ เซอร์วิส" in text or "มนต์โลจิสติกส์เซอร์วิส" in text:
        result["tax_id"] = MON_LOGISTICS_TAX_ID
    else:
        # Extract Tax ID
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
    
    # Extract Branch
    branch_config = common_fields_config.get("branch", {})
    default_hq = branch_config.get("default_hq", "00000")
    pad_zeros = branch_config.get("pad_zeros", 5)
    
    # Priority 1: Look for checked checkbox with branch (☑ สาขาที่ X)
    checked_branch_match = re.search(r'[☑✓✔]\s*สาขา(?:ที่)?\s*(\d+)', text)
    if checked_branch_match:
        result["branch"] = checked_branch_match.group(1).zfill(pad_zeros)
    else:
        # Priority 2: Look for checked Head Office checkbox
        checked_hq_match = re.search(r'[☑✓✔]\s*(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office)', text, re.IGNORECASE)
        if checked_hq_match:
            result["branch"] = default_hq
        else:
            # Priority 3: Look for "สาขาที่ออกใบกำกับภาษี คือ XXXXX" pattern (vendor's branch)
            vendor_branch_match = re.search(r'สาขาที่ออกใบกำกับภาษี\s*(?:คือ|:)?\s*(\d{1,5})', text, re.IGNORECASE)
            if vendor_branch_match:
                result["branch"] = vendor_branch_match.group(1).zfill(pad_zeros)
            else:
                # Priority 4: Look for "สำนักงานใหญ่ XXXXX" or "HEAD OFFICE XXXXX" (head office with number)
                hq_with_num_match = re.search(r'(?:สำนักงานใหญ่|HEAD\s*OFFICE)\s*[:\s]?\s*(\d{5})', text, re.IGNORECASE)
                if hq_with_num_match:
                    result["branch"] = hq_with_num_match.group(1).zfill(pad_zeros)
                else:
                    # Priority 5: Standard branch pattern (without checkbox context)
                    branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})(?!\d)", text, re.IGNORECASE)
                    if branch_match:
                        result["branch"] = branch_match.group(1).zfill(pad_zeros)
                    else:
                        # Priority 6: Fall back to Head Office keywords (without number)
                        ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
                        if ho_match:
                            result["branch"] = default_hq
                        else:
                            # Default to Head Office if nothing found
                            result["branch"] = default_hq
    
    return result


def parse_ocr_data(text, templates, doc_type="auto"):
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
    
    if not text or not templates:
        return result
    
    # Detect document type
    if doc_type == "auto":
        detected_type = detect_document_type(text, templates)
    else:
        detected_type = doc_type if doc_type in templates.get("templates", {}) else "invoice"
    
    result["document_type"] = detected_type
    
    template = templates.get("templates", {}).get(detected_type, {})
    result["document_type_name"] = template.get("name", detected_type)
    
    # Extract common fields
    common_fields = templates.get("common_fields", {})
    common_result = extract_common_fields(text, common_fields)
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
        
        text_to_search = text
        skip_lines = field_config.get("skip_lines", 0)
        if skip_lines > 0:
            lines = text.split('\n')
            text_to_search = '\n'.join(lines[skip_lines:])
        
        value = extract_field_by_patterns(text_to_search, patterns, options)
        
        # Fallback for amount
        if not value and field_config.get("fallback") == "last_amount":
            amounts = re.findall(r"([\d,]+\.\d{2})", text)
            value = amounts[-1] if amounts else ""
        
        # Reject 13-digit numbers for document_no
        if field_name == "document_no" and value:
            digits_only = re.sub(r'\D', '', value)
            if len(digits_only) == 13 and digits_only.isdigit():
                value = ""
        
        # Fallback for date: search for common date formats if not found
        if field_name == "date" and not value:
            # Search for date patterns: xx/xx/xxxx, xx-xx-xxxx, xx.xx.xxxx, xx/xx/xx, xx-xx-xx
            date_patterns = [
                r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})',  # dd/mm/yyyy or dd-mm-yyyy
                r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2})'   # dd/mm/yy or dd-mm-yy
            ]
            for date_pattern in date_patterns:
                date_matches = re.findall(date_pattern, text)
                if date_matches:
                    value = date_matches[0]
                    break
        
        # Normalize date format: convert dashes/dots to slashes
        if field_name == "date" and value:
            value = re.sub(r'[-.]', '/', value)
        
        if field_name in ["document_no", "date", "amount"]:
            result[field_name] = value
        else:
            result["extra_fields"][field_name] = value
    
    return result


def main():
    print("=" * 60)
    print("Test Regex Extraction Script")
    print("=" * 60)
    print(f"Example Doc Dir: {EXAMPLE_DOC_DIR}")
    print(f"Templates File: {TEMPLATES_FILE}")
    print(f"Output File: {OUTPUT_FILE}")
    print()
    
    # Load templates
    templates = load_templates()
    if not templates:
        return
    print(f"Loaded templates: {list(templates.get('templates', {}).keys())}")
    
    # Load vendor master
    vendor_df = load_vendor_master()
    if vendor_df is not None:
        print(f"Loaded {len(vendor_df)} vendor records")
    
    # Get all .txt files
    txt_files = sorted([f for f in os.listdir(EXAMPLE_DOC_DIR) if f.endswith('.txt')])
    print(f"\nFound {len(txt_files)} .txt files")
    print()
    
    data_rows = []
    
    for filename in txt_files:
        file_path = os.path.join(EXAMPLE_DOC_DIR, filename)
        
        # Extract page number from filename
        page_match = re.search(r'_page(\d+)\.txt$', filename)
        page_num = int(page_match.group(1)) if page_match else 0
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            print(f"Error reading {filename}: {e}")
            continue
        
        # Parse content
        parsed = parse_ocr_data(content, templates)
        
        print(f"  {filename}: Type={parsed['document_type_name']}, TaxID={parsed['tax_id']}, Branch={parsed['branch']}, DocNo={parsed['document_no']}")
        
        # Create hyperlink formula pointing to the source PDF file
        # Extract base filename without _pageX suffix to get PDF name
        base_name_match = re.match(r'^(.+?)_page\d+\.txt$', filename)
        if base_name_match:
            pdf_name = base_name_match.group(1) + ".pdf"
        else:
            # Fallback: just replace .txt with .pdf
            pdf_name = filename.replace('.txt', '.pdf')
        
        pdf_path = os.path.join(EXAMPLE_DOC_DIR, pdf_name)
        hyperlink_formula = f'=HYPERLINK("{pdf_path}", "{filename}")'

        
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
        
        # Handle CY INSTRUCTION
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
    
    if not data_rows:
        print("No data extracted.")
        return
    
    # Create DataFrame
    df = pd.DataFrame(data_rows)
    
    # Remap Sahatthai Invoice to ใบกำกับภาษี/Invoice for Summary Excel
    if 'Document Type' in df.columns:
        df['Document Type'] = df['Document Type'].replace({
            'Sahatthai Invoice': 'ใบกำกับภาษี/Invoice',
            'sahatthai_invoice': 'ใบกำกับภาษี/Invoice'
        })
    
    # Merge with vendor data
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
        df.rename(columns={'Vendor code SAP': 'Vendor code', 'ชื่อบริษัท': 'Vendor Name'}, inplace=True)
        df.drop(columns=['เลขประจำตัวผู้เสียภาษี', 'สาขา'], inplace=True, errors='ignore')
    else:
        df['Vendor code'] = ""
        df['Vendor Name'] = ""
    
    # Reorder columns - put important ones first (matching Extract_Inv.py)
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
    
    # Save to Excel with multiple sheets
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
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
            
            # Define column configuration for each sheet (matching Extract_Inv.py)
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
                final_cols = [col for col in target_cols if col in group_df.columns]
                
                # Final filtering and reordering
                group_df = group_df[final_cols]
                
                # Reset index for clean output
                group_df = group_df.reset_index(drop=True)
                group_df.to_excel(writer, index=False, sheet_name=str(sheet_name))
                print(f"   -> Sheet '{sheet_name}': {len(group_df)} rows")
        
        print(f"\nSuccess! Output saved at: {OUTPUT_FILE}")
        print(f"Total rows: {len(df)}")
        
    except Exception as e:
        print(f"Error saving Excel: {e}")


if __name__ == "__main__":
    main()
