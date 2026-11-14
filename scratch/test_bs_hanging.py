import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import report_generator_python, pymupdf

def test_bs(cust_name):
    doc = docx.Document('DOC/Belt Scale.docx')
    p3 = doc.paragraphs[3]
    p4 = doc.paragraphs[4]
    
    def set_hanging_indent_tab(p, text_left, text_right, tab_dxa=4200, font_pt=11.0):
        p.text = ''
        pPr = p._p.get_or_add_pPr()
        for old_tabs in pPr.findall(qn('w:tabs')):
            pPr.remove(old_tabs)
        tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_dxa}"/></w:tabs>')
        pPr.append(tabs_xml)
        
        for old_ind in pPr.findall(qn('w:ind')):
            pPr.remove(old_ind)
        ind_xml = parse_xml(f'<w:ind {nsdecls("w")} w:left="{tab_dxa}" w:firstLine="-{tab_dxa}"/>')
        pPr.append(ind_xml)
        
        r_l = p.add_run(f'{text_left}\t')
        r_l.bold = True
        r_l.font.size = Pt(font_pt)
        r_l.font.name = 'Times New Roman'
        r_r = p.add_run(text_right)
        r_r.bold = True
        r_r.font.size = Pt(font_pt)
        r_r.font.name = 'Times New Roman'

    r_len = len(cust_name)
    if r_len > 45:
        tab_dxa = 3600
        font_pt = 9.5
    else:
        tab_dxa = 4200
        font_pt = 11.0

    set_hanging_indent_tab(p3, 'JOB No: SY1529', f'CUSTOMER NAME: {cust_name}', tab_dxa, font_pt)
    set_hanging_indent_tab(p4, 'CAPACITY: 500 TPH', 'CONVEYOR NO: BC-01', tab_dxa, font_pt)

    out_docx = f'scratch/test_bs_{cust_name[:10]}.docx'
    out_pdf = f'scratch/test_bs_{cust_name[:10]}.pdf'
    doc.save(out_docx)
    report_generator_python.convert_docx_to_pdf(out_docx, out_pdf)
    doc_pdf = pymupdf.open(out_pdf)
    print(f'Cust: {cust_name[:20]} | Pages: {len(doc_pdf)}')
    p0 = doc_pdf[0]
    for block in p0.get_text('dict')['blocks']:
        if block.get('type') == 0:
            for line in block.get('lines', []):
                line_text = ''.join(s.get('text', '') for s in line.get('spans', []))
                y = line['bbox'][1]
                if 80 < y < 130:
                    print(f'  y={y:.1f}:')
                    for sp in line['spans']:
                        print(f"     x={sp['bbox'][0]:.1f} size={sp['size']:.1f} font={sp['font']}: {repr(sp['text'])}")

test_bs('Exceltech Thermal Optimizer Private Limited')
test_bs('ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD')
