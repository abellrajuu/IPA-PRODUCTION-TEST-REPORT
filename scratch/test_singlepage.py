import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import report_generator_python as rgp
import fitz

d = docx.Document('DOC/Belt Scale.docx')
p3 = d.paragraphs[3]
p4 = d.paragraphs[4]

tab_dxa = 2800
def set_header_line(p, left_t, right_t, font_sz=10.0):
    p.text = ''
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')):
        pPr.remove(old_tabs)
    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_dxa}"/></w:tabs>')
    pPr.append(tabs_xml)
    r1 = p.add_run(f'{left_t}\t')
    r1.font.size = Pt(11)
    r1.bold = True
    r2 = p.add_run(right_t)
    r2.font.size = Pt(font_sz)
    r2.bold = True

long_cust = 'ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD'
set_header_line(p3, 'JOB No: JOB-SY1615', f'CUSTOMER NAME: {long_cust}', 10.0)
set_header_line(p4, 'CAPACITY: 1200 TPH', 'CONVEYOR NO: CV-101', 11.0)

d.save('tests/test_bs_singlepage.docx')
pdf = rgp.convert_docx_to_pdf('tests/test_bs_singlepage.docx')
doc = fitz.open(pdf)
print('New PDF Page Count:', len(doc))
text = doc[0].get_text()
for line in text.split('\n')[:12]:
    print('  ', repr(line))
