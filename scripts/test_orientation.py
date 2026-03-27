"""测试 A4 纵向/横向页面设置"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn

for orient, label in [("纵", "A4纵向"), ("横", "A4横向")]:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width  = int(21 / 2.54 * 1440)
    sec.page_height = int(29.7 / 2.54 * 1440)
    if orient == "横":
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.top_margin = int(1.1 / 2.54 * 1440)
    sec.bottom_margin = int(1.2 / 2.54 * 1440)
    sec.left_margin = int(2.4 / 2.54 * 1440)
    sec.right_margin = int(1.7 / 2.54 * 1440)
    p = doc.add_paragraph(f"{label} 测试")
    doc.save(os.path.join(os.environ['TEMP'], f'test_{orient}.docx'))
    print(f"{label} OK")

print("All orientation tests passed")
