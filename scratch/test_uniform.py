import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import report_generator_python, pymupdf

doc = docx.Document('DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM mv calibration.docx')

def set_para_with_tab(p, text_left, text_right, tab_pos_dxa=4200, font_size_pt=11.0, bold=None):
    p.text = ''
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')):
        pPr.remove(old_tabs)
    for old_ind in pPr.findall(qn('w:ind')):
        pPr.remove(old_ind)
    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_pos_dxa}"/></w:tabs>')
    pPr.append(tabs_xml)
    ind_xml = parse_xml(f'<w:ind {nsdecls("w")} w:left="{tab_pos_dxa}" w:firstLine="-{tab_pos_dxa}"/>')
    pPr.append(ind_xml)
    r_l = p.add_run(f'{text_left}\t')
    if bold is not None: r_l.bold = bold
    r_l.font.size = Pt(font_size_pt)
    r_l.font.name = 'Times New Roman'
    r_r = p.add_run(text_right)
    if bold is not None: r_r.bold = bold
    r_r.font.size = Pt(font_size_pt)
    r_r.font.name = 'Times New Roman'

customer = 'Exceltech Thermal Optimizer Private Limited'
if len(customer) > 55: hdr_tab, hdr_pt = 3600, 9.5
elif len(customer) > 40: hdr_tab, hdr_pt = 3900, 10.0
else: hdr_tab, hdr_pt = 4200, 11.0

# P2
doc.paragraphs[2].text = '        SYSTEM  : DIGITAL INDICATOR'
for r in doc.paragraphs[2].runs:
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Times New Roman'

# P3 & P4 with identical tab and font
set_para_with_tab(doc.paragraphs[3], '        JOB No.    : JOB-SY1550', f'CUSTOMER NAME: {customer}', hdr_tab, hdr_pt, bold=True)
set_para_with_tab(doc.paragraphs[4], '        CAPACITY: 100KG', 'TAG. NO: ', hdr_tab, hdr_pt, bold=True)

doc.save('scratch/test_uniform.docx')
report_generator_python.convert_docx_to_pdf('scratch/test_uniform.docx', 'scratch/test_uniform.pdf')

doc_pdf = pymupdf.open('scratch/test_uniform.pdf')
p0 = doc_pdf[0]
for block in p0.get_text('dict')['blocks']:
    if block.get('type') == 0:
        for line in block.get('lines', []):
            line_text = ''.join(s.get('text', '') for s in line.get('spans', []))
            y = line['bbox'][1]
            if 60 < y < 115:
                print(f'Line y={y:.1f}:')
                for sp in line['spans']:
                    print(f"   x={sp['bbox'][0]:.1f} size={sp['size']:.1f} font={sp['font']}: {repr(sp['text'])}")
