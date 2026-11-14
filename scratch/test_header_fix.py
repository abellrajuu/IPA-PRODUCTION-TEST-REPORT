import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import report_generator_python, pymupdf

doc = docx.Document('DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM mv calibration.docx')

# Keep P1 empty
doc.paragraphs[1].text = ''

# P2: SYSTEM: DIGITAL INDICATOR
doc.paragraphs[2].text = '        SYSTEM  : DIGITAL INDICATOR'
for r in doc.paragraphs[2].runs:
    r.font.size = Pt(11)
    r.bold = True
    r.font.name = 'Times New Roman'

# Test with tab stop at 4400 dxa (220 pt)
def set_header_tab(p, text_left, text_right, tab_dxa=4000, font_pt=10.0):
    p.text = ''
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')):
        pPr.remove(old_tabs)
    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_dxa}"/></w:tabs>')
    pPr.append(tabs_xml)
    r_l = p.add_run(f'{text_left}\t')
    r_l.bold = True
    r_l.font.size = Pt(font_pt)
    r_l.font.name = 'Times New Roman'
    r_r = p.add_run(text_right)
    r_r.bold = True
    r_r.font.size = Pt(font_pt)
    r_r.font.name = 'Times New Roman'

set_header_tab(doc.paragraphs[3], '        JOB No.    : SY1529', 'CUSTOMER: ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD', 3800, 9.5)
set_header_tab(doc.paragraphs[4], '        CAPACITY: 100KG', 'TAG. NO: ', 3800, 9.5)

doc.save('scratch/test_sy1550_tab.docx')
report_generator_python.convert_docx_to_pdf('scratch/test_sy1550_tab.docx', 'scratch/test_sy1550_tab.pdf')

doc_pdf = pymupdf.open('scratch/test_sy1550_tab.pdf')
p0 = doc_pdf[0]
for block in p0.get_text('dict')['blocks']:
    if block.get('type') == 0:
        for line in block.get('lines', []):
            line_text = ''.join(s.get('text', '') for s in line.get('spans', []))
            y = line['bbox'][1]
            if 50 < y < 140:
                print(f'Line y={y:.1f}:')
                for sp in line['spans']:
                    print(f"   x={sp['bbox'][0]:.1f} size={sp['size']:.1f} font={sp['font']}: {repr(sp['text'])}")
