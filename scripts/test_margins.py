"""Test section margin and page size settings"""
from docx import Document
from docx.shared import Cm, Inches

for orient, label in [("纵", "纵向"), ("横", "横向")]:
    doc = Document()
    sec = doc.sections[0]

    if orient == "横":
        sec.page_width  = int(29.7 / 2.54 * 1440)
        sec.page_height = int(21 / 2.54 * 1440)
    else:
        sec.page_width  = int(21 / 2.54 * 1440)
        sec.page_height = int(29.7 / 2.54 * 1440)

    sec.top_margin    = Cm(1.1)
    sec.bottom_margin = Cm(1.2)
    sec.left_margin   = Cm(2.4)
    sec.right_margin  = Cm(1.7)

    doc.add_paragraph(f"A4{label} 上1.1下1.2左2.4右1.7cm")

    from docx.oxml.ns import qn
    pg_sz = sec._sectPr.find(qn('w:pgSz'))
    pg_mar = sec._sectPr.find(qn('w:pgMar'))
    w = int(pg_sz.get(qn('w:w'))) / 1440 * 2.54
    h = int(pg_sz.get(qn('w:h'))) / 1440 * 2.54
    top = int(pg_mar.get(qn('w:top'))) / 1440 * 2.54
    bot = int(pg_mar.get(qn('w:bottom'))) / 1440 * 2.54
    lft = int(pg_mar.get(qn('w:left'))) / 1440 * 2.54
    rgt = int(pg_mar.get(qn('w:right'))) / 1440 * 2.54
    print(f"{label}: 页面={w:.1f}cm×{h:.1f}cm 上={top:.2f} 下={bot:.2f} 左={lft:.2f} 右={rgt:.2f}")

    doc.save(f"C:/Users/ADMINI~1/AppData/Local/Temp/margin_{orient}.docx")

print("Done")
