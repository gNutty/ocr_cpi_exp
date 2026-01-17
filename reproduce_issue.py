import re

def extract_common_fields(text):
    """Extract common fields (tax_id, branch) that apply to all document types"""
    result = {"branch": ""}
    
    default_hq = "00000"
    pad_zeros = 5
    
    # Check for Head Office keywords first (ORIGINAL LOGIC)
    ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
    if ho_match:
        print(f"Matched Head Office: {ho_match.group(0)}")
        result["branch"] = default_hq
    else:
        # Try to find branch number
        branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})", text, re.IGNORECASE)
        if branch_match:
            print(f"Matched Branch Number: {branch_match.group(1)}")
            result["branch"] = branch_match.group(1).zfill(pad_zeros)
    
    return result

def extract_common_fields_fixed(text):
    """Refined logic: Check for specific branch number FIRST"""
    result = {"branch": ""}
    
    default_hq = "00000"
    pad_zeros = 5
    
    # Try to find branch number FIRST
    branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})", text, re.IGNORECASE)
    if branch_match:
        print(f"Matched Branch Number: {branch_match.group(1)}")
        result["branch"] = branch_match.group(1).zfill(pad_zeros)
    else:
        # Check for Head Office keywords
        ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
        if ho_match:
            print(f"Matched Head Office: {ho_match.group(0)}")
            result["branch"] = default_hq
            
    return result

# Text from example_doc/all vendorP15_page2.txt
text_content = """G-FORTUNE
บริษัท เกรทติ้ง ฟอร์จูน คอนเทนเนอร์ เซอร์วิส (ประเทศไทย) จำกัด
สำนักงานใหญ่ : 333/4 หมู่ 16 ต.บางแก้ว อ.บางพลี จ.สมุทรปราการ 10540 โทร. 0-2316-7288
Head Office : 333/4 Moo 16, Tambon Bang Kaeo, Amphur Bangplee, Samutprakarn 10540 Tel. 0-2316-7288
สาขาที่ออกใบกำกับภาษี :
☑ สาขาที่ 00006 : 42/59-42/60 Moo 3, Racha Thewa , Bangplee , Samutprakarn 10540 Tel. 0-2316-7288
Branch 00006 : 42/59-42/60 Moo 3, Racha Thewa , Bangplee , Samutprakarn 10540 Tel. 0-2316-7288
"""

print("--- Testing Original Logic ---")
res = extract_common_fields(text_content)
print(f"Result: {res['branch']}")
print(f"Expected: 00006")

print("\n--- Testing Fixed Logic ---")
res_fixed = extract_common_fields_fixed(text_content)
print(f"Result: {res_fixed['branch']}")
print(f"Expected: 00006")
