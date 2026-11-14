import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import report_generator_python, pymupdf

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

    r_left = p.add_run(f'{text_left}\t')
    if bold is not None:
        r_left.bold = bold
    r_left.font.size = Pt(font_size_pt)
    r_left.font.name = 'Times New Roman'

    r_right = p.add_run(text_right)
    if bold is not None:
        r_right.bold = bold
    r_right.font.size = Pt(font_size_pt)
    r_right.font.name = 'Times New Roman'

doc = docx.Document('DOC/PP-05_08 Crane Scale.docx')
cust = 'ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD'
set_para_with_tab(doc.paragraphs[2], '        JOB No.    : SY1615', f'CUSTOMER: {cust}', 3600, 9.5, bold=True)
set_para_with_tab(doc.paragraphs[3], '        CAPACITY  : 10 Tonnes', 'STAMP NO: STAMP-2026-01', 3600, 9.5, bold=True)
set_para_with_tab(doc.paragraphs[4], '        Load cell. Model No. : LC-CWS-10T', 'Serial No. : SN-LC-001', 3600, 9.5, bold=True)

out_docx = 'scratch/test_cs_electromech.docx'
out_pdf = 'scratch/test_cs_electromech.pdf'
doc.save(out_docx)
report_generator_python.convert_docx_to_pdf(out_docx, out_pdf)
doc_pdf = pymupdf.open(out_pdf)
print('Pages:', len(doc_pdf))
for b in doc_pdf[0].get_text('dict')['blocks']:
    if b.get('type') == 0:
        for l in b['lines']:
            text = ''.join(s['text'] for s in l['spans'])
            if 'ELECTRO' in text or 'INDIA' in text or 'JOB No' in text:
                print(f'Line y={l["bbox"][1]:.1f}:')
                for s in l['spans']:
                    print(f'  x={s["bbox"][0]:.1f} font={s["font"]} size={s["size"]:.1f} text={repr(s["text"])}')
