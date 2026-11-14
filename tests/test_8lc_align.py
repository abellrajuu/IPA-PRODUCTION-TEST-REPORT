import os
import sys
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import pymupdf as fitz

BASE_DIR = r"C:\Users\Admin\Desktop\PDN TEST REPORT"
sys.path.insert(0, BASE_DIR)
import report_generator_python as rg

test_docx = os.path.join(BASE_DIR, "tests", "test_bs8_tab.docx")
test_pdf = os.path.join(BASE_DIR, "tests", "test_bs8_tab.pdf")

d = docx.Document(os.path.join(BASE_DIR, "backend_docx_templates", "PP-05_03 Belt Scale - 8 LC.docx"))
t0 = d.tables[0]
col_w = [2592, 3744, 2304, 2736]
for r in t0.rows:
    for c_i, w in enumerate(col_w):
        r.cells[c_i].width = int(w * 635)
        r.cells[c_i]._tc.tcPr.tcW.w = w

if hasattr(t0._tbl, 'tblGrid') and t0._tbl.tblGrid is not None:
    for c_i, w in enumerate(col_w):
        t0._tbl.tblGrid.gridCol_lst[c_i].w = w

pairs = [("101", "505"), ("2024-55", "2024-66"), ("LC-CAN-03", "LC-CAN-07"), ("9", "8")]

c1 = t0.rows[1].cells[1]
for idx, (s_left, s_right) in enumerate(pairs):
    p = c1.paragraphs[4 + idx]
    p.text = f"{idx+1}) {s_left}\t{idx+5}) {s_right}"
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')):
        pPr.remove(old_tabs)
    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1750"/></w:tabs>')
    pPr.append(tabs_xml)
    for r in p.runs:
        r.font.size = Pt(11)

d.save(test_docx)
rg.convert_docx_to_pdf(test_docx, test_pdf)

doc = fitz.open(test_pdf)
print("Page count with tab stop:", len(doc))
doc[0].get_pixmap(dpi=150).save(os.path.join(BASE_DIR, "tests", "test_bs8_tab_p0.png"))
print("Done saving test_bs8_tab_p0.png")
