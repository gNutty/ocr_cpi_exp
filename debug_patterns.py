import re

text = """บริษัท สหไทย เทอร์มินอล จำกัด (มหาชน)
สำนักงานใหญ่ 00000 เลขที่ 5/1 ม.3 ถนนปู่เจ้าสมชาย ต.บางกอกใหญ่ อ.พระประแดง จ.สมุทรปราการ 10130
สาขา 00001 เลขที่ 79 หมู่ 3 ถนนปู่เจ้าสมชาย ต.บางกอกใหญ่ อ.พระประแดง จ.สมุทรปราการ 10130
Tel. +662 386-8000 Fax. +662 754-4141
เลขประจำตัวผู้เสียภาษีอากร 0107560000162
SAHATTHAI TERMINAL PUBLIC COMPANY LIMITED
HEAD OFFICE 00000 : 5/1 Moo 3 Pochasamngrai Rd., Bangyaphrak Phrapradaeng Samutprakan 10130
BRANCH 00001 : 79 Moo 3 Pochasamngrai Rd., Bangyaphrak Phrapradaeng Samutprakan 10130
BRANCH 00002 : 2105 Moo 12 Thong Sukha, Sri-Racha Chonburi 20230
Tel. +662 386-8000 Fax. +662 754-4141
ORIGINAL /อันฉบับ
☐ อัตราภาษีร้อยละ 7 ☐ อัตราภาษีศูนย์
วันที่ / Date : 12/09/2025
เลขที่ / NO. G250903985
ได้รับ จาก/Received From C.P.INTERTRADE CO.,LTD."""

patterns = [
    r"(BK\d+-OUT\s*\d{5,6})",
    r"No\.?\s*[:\.]?\s*(LF-\d+)",
    r"เลขที่\s*/\s*NO\.?\s*([A-Za-z]\d{5,})",
    r"เลขที่\s+([A-Za-z]\d{6,})",
    r"เลขที่\s*NO\.?\s*[/ \s]*(_?[A-Za-z]\d{5,})",
    r"Tax\s*Invoice\s*No[\s\.:]*([A-Za-z0-9\-\/]*\d+[A-Za-z0-9\-\/]*)",
    r"Tax\s*Invoice\s*No\.?\s*[:\.]?\s*([A-Za-z0-9\-\/]*\d+[A-Za-z0-9\-\/]*)",
    r"RECEIPT\s*No\.?\s*[:\.]?\s*([A-Za-z0-9\-\/]*\d+[A-Za-z0-9\-\/]*)",
    r"เลขที่/NO\.?\s*[:\.]?\s*([A-Za-z0-9\-\/]*\d+[A-Za-z0-9\-\/]*)",
    r"เลขที่\s*[:\.]?\s*(\d{2}\s*-\s*\d{6}\s*-\s*\d{4})",
    r"No\.?\s*[:\.]?\s*([A-Z]+-[\d]+)",
    r"Invoice\s*No\.?\s*[:\.]?\s*([A-Za-z0-9\-\/]*\d+[A-Za-z0-9\-\/]*)",
    r"เลขที่\s*[:\.]?\s*([A-Za-z0-9\-\/]*\d+[A-Za-z0-9\-\/]*)",
    r"No\.?\s*[:\.]?\s*([A-Za-z0-9\-\/]*\d[A-Za-z0-9\-\/]*)"
]

print(f"Text Lines: {len(text.splitlines())}")

for i, pattern in enumerate(patterns):
    matches = re.finditer(pattern, text, re.IGNORECASE)
    for match in matches:
        value = match.group(1) if match.lastindex and match.lastindex >= 1 else match.group(0)
        start = match.start()
        # Find which line it is in
        line_num = text.count('\n', 0, start) + 1
        
        print(f"Pattern {i} matched at Line {line_num}:")
        print(f"  Full Match: '{match.group(0)}'")
        print(f"  Extracted Group: '{value}'")
        if value == "5/1":
            print("  *** FOUND TARGET VALUE '5/1' ***")
