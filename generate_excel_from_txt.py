import os
import pandas as pd
import Extract_Inv  # Import the module with corrected logic

# Define paths
SOURCE_DIR = r"d:\Project\ocr\ocr_cpi_exp\example_doc"
OUTPUT_DIR = r"d:\Project\ocr\output"
TXT_FILE = "all vendor_page2.txt"

def generate_excel():
    print(f"--- Generating Excel from {TXT_FILE} ---")
    
    # Check file exists
    txt_path = os.path.join(SOURCE_DIR, TXT_FILE)
    if not os.path.exists(txt_path):
        print(f"Error: {txt_path} not found.")
        return

    # Read text
    with open(txt_path, 'r', encoding='utf-8') as f:
        text = f.read()
    
    # Load templates
    templates = Extract_Inv.load_templates()
    
    # Parse text using the Import (which has FIXED logic)
    parsed = Extract_Inv.parse_ocr_data_with_template(text, templates, doc_type="cy_instruction")
    
    print(f"Detected: {parsed['document_type_name']}")
    
    # Build Row Data (Mirroring Extract_Inv.py logic)
    row_data = {
        "Link PDF": "N/A",
        "Page": 2,
        "Document Type": parsed["document_type_name"],
        "VendorID_OCR": parsed["tax_id"],
        "Branch_OCR": parsed["branch"],
        "Document No": parsed["document_no"],
        "Date": parsed["date"],
        "Amount": parsed["amount"],
    }
    
    extra = parsed.get("extra_fields", {})
    
    # Logic for CyBooking (Now fixed in Extract_Inv.py)
    # Note: We rely on parse_ocr_data_with_template from Extract_Inv.py
    # But wait, Extract_Inv.py returns 'cy_booking' in extra_fields.
    # The concatenation logic was in the MAIN function of Extract_Inv.py, NOT in parse_ocr_data_with_template
    # So we must REPLICATE the FIXED logic here.
    
    booking_no = extra.get("cy_booking", "")
    # FIXED LOGIC: Just use booking_no
    cy_booking = booking_no
    
    row_data["CyOrg"] = extra.get("cy_org", "")
    row_data["CyExporter"] = extra.get("cy_exporter", "")
    row_data["CyInvoiceNo"] = extra.get("cy_invoice_no", "")
    row_data["CyBooking"] = cy_booking
    row_data["CyQty"] = extra.get("cy_qty", "")
    
    print(f"CyBooking Value: '{cy_booking}'")
    
    # Create DataFrame
    df = pd.DataFrame([row_data])
    
    # Load Vendor Master for mapping (optional but good for completeness)
    vendor_df = Extract_Inv.load_vendor_master()
    if vendor_df is not None:
         # Clean branch
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
    
    # Output path
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    output_excel_path = os.path.join(OUTPUT_DIR, "summary_ocr.xlsx")
    
    try:
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
             df.to_excel(writer, index=False, sheet_name='CY_INSTRUCTION')
        print(f"Success! Saved to {output_excel_path}")
    except Exception as e:
        print(f"Error saving excel: {e}")

if __name__ == "__main__":
    generate_excel()
