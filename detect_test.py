import json
import os

TEMPLATES_FILE = "document_templates.json"

def load_templates():
    with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def detect_document_type(text, templates):
    """Current logic from Extract_Inv.py"""
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
    # Strict Header: First 4 lines
    strict_header = "\n".join(lines[:4]).lower()
    # General Header: First 15 lines
    header_text = "\n".join(lines[:15]).lower()
    
    scores = {}
    
    for doc_type, template in templates.get("templates", {}).items():
        if doc_type in priority_types:
            continue
        
        keywords = template.get("detect_keywords", [])
        score = 0
        for keyword in keywords:
            kw_low = keyword.lower()
            
            # Top 4 lines get SUPER priority
            if kw_low in strict_header:
                score += 50
            
            # Normal Header match
            elif kw_low in header_text:
                score += 10
            
            # Body match
            if kw_low in text_lower:
                score += 1
                
        if score > 0:
            scores[doc_type] = score
            print(f"Type: {doc_type}, Score: {score}")
    
    if scores:
        return max(scores, key=scores.get)
    return "invoice"

text_page1 = """บริษัท ไดนามิคทรานสปอร์ต จำกัด
เลขที่ 3 อาคาร ซี.พี.ทาวเวอร์2(ฟอร์จูนทาวน์) ชั้น19 โซน C,D2 ถนนรัชดาภิเษก แขวงดินแดง เขตดินแดง กรุงเทพฯ 10400
เลขประจำตัวผู้เสียภาษี:0105532115191
ใบวางบิล
เลขที่ 6800707435
วันที่ 27/08/2568
รหัสลูกค้า 1400180
เลขประจำตัวผู้เสียภาษี
ผู้ว่าจ้าง บริษัท ซี.พี.อินเตอร์เทรด จำกัด นครหลวง
ที่อยู่ 313 อาคารซี.พี.ทาวเวอร์ ถ.สีลม แขวงสีลม เขตบางรัก กรุงเทพฯ 10500
จุดวางบิล วันครบกำหนดชำระ 26/09/2568
อาคารฟอร์จูนทาวน์
 ลำดับ ประเภทเอกสาร เลขที่เอกสาร วันที่ จำนวนเงิน 1 ใบแจ้งหนี้อื่นๆ 680162010028 23/08/2568 1,720.80 2 ใบแจ้งหนี้อื่นๆ 680162010029 23/08/2568 1,720.80 3 ใบแจ้งหนี้อื่นๆ 680162010030 23/08/2568 1,720.80 4 ใบแจ้งหนี้อื่นๆ 680162010031 23/08/2568 1,720.80 5 ใบแจ้งหนี้อื่นๆ 680162010032 23/08/2568 1,720.80 6 ใบแจ้งหนี้อื่นๆ 680162010033 23/08/2568 1,720.80 7 ใบแจ้งหนี้อื่นๆ 680162010034 23/08/2568 1,426.00 8 ใบแจ้งหนี้อื่นๆ 680162010035 23/08/2568 1,426.00 9 ใบแจ้งหนี้อื่นๆ 680162010036 23/08/2568 1,426.00 10 ใบแจ้งหนี้อื่นๆ 680162010037 23/08/2568 1,426.00 11 ใบแจ้งหนี้อื่นๆ 680162010038 23/08/2568 1,426.00 12 ใบแจ้งหนี้อื่นๆ 680162010039 23/08/2568 1,426.00 13 ใบแจ้งหนี้อื่นๆ 680162010040 23/08/2568 1,426.00 14 ใบแจ้งหนี้อื่นๆ 680162010041 23/08/2568 1,426.00 15 ใบแจ้งหนี้อื่นๆ 680162010042 23/08/2568 1,426.00 16 ใบแจ้งหนี้อื่นๆ 680162010043 23/08/2568 1,426.00 17 ใบแจ้งหนี้อื่นๆ 680162010044 23/08/2568 1,426.00 18 ใบแจ้งหนี้อื่นๆ 680162010045 23/08/2568 1,426.00 รวมเงิน 27,436.80 
(signature)
ผู้วางบิล
(signature)
ผู้รับบิล
ต้นฉบับ"""

templates = load_templates()
print("--- Check Classification for Vendor Page 1 ---")
result = detect_document_type(text_page1, templates)
print(f"Detected Type: {result}")
