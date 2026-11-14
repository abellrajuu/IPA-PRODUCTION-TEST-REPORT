import docx, os, sys
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from report_generator_python import convert_docx_to_pdf
import pymupdf

def test_cws():
    doc = docx.Document('backend_docx_templates/PP-05_08 Crane Scale.docx')

    def set_para_tab(p, left_text, right_text, tab_pos=5000, font_size=10, bold_label=False):
        if not p: return
        p.text = f"{left_text}\t{right_text}"
        pPr = p._p.get_or_add_pPr()
        for old_tabs in pPr.findall(qn('w:tabs')):
            pPr.remove(old_tabs)
        tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_pos}"/></w:tabs>')
        pPr.append(tabs_xml)
        for r in p.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(font_size)
            r.bold = bold_label

    set_para_tab(doc.paragraphs[2], 'JOB No.\t: JOB-CWS-2026-101', 'CUSTOMER: UltraTech Cement Works', 5000, 10, False)
    set_para_tab(doc.paragraphs[3], 'CAPACITY  : 10 Tonnes', 'STAMP NO: STAMP-2026-88', 5000, 10, False)
    set_para_tab(doc.paragraphs[4], 'Load cell. Model No. : LC-CWS-10T', 'Serial No. : SN-LC-9901', 5000, 10, False)
    doc.paragraphs[8].text = '\t   Resolution :\tL: 2.0 kg/tonnes'

    t0 = doc.tables[0]
    r1 = t0.rows[1]
    r1.cells[0].text = 'Model: CS-10T\nSl. No: SN-CS-101'
    r1.cells[1].text = 'BC-12V\nSl. No: CH-881'
    r1.cells[2].text = 'RDU-RF-433\nSl. No: RD-552'
    r1.cells[3].text = 'Sl. No: HK-901'
    r1.cells[4].text = 'Sl. No: SH-902'

    t1 = doc.tables[1]
    t1.rows[1].cells[3].text = '6.52 V (PASS)'
    t1.rows[2].cells[3].text = 'Checked & OK'
    t1.rows[3].cells[3].text = 'Working Normal'

    doc.paragraphs[13].text = '        When the battery connected: 6.8 V (Connected & Charging)'
    doc.paragraphs[15].text = 'Low battery voltage set (CS): 5.8 V (Cut-off Alarm Verified)'
    doc.paragraphs[17].text = 'Check for the remote operation: Verified up to 25 meters (No packet loss)'
    doc.paragraphs[29].text = '                 Instrument Used: Fluke 179 True RMS Multimeter (Cal. Valid)'

    set_para_tab(doc.paragraphs[31], 'Checked by: Quality Engineer', 'Approved by: HOD-PDN', 6500, 10, False)
    set_para_tab(doc.paragraphs[32], 'Date: 08-Oct-2026', '(HOD - PDN)', 6500, 10, False)

    out_docx = 'tests/test_cs_styled.docx'
    doc.save(out_docx)
    convert_docx_to_pdf(out_docx)

    pdf = pymupdf.open('tests/test_cs_styled.pdf')
    print('Page count:', len(pdf))
    assert len(pdf) == 1, 'Error: Must be 1 page!'
    page = pdf[0]
    page.get_pixmap(dpi=150).save('C:/Users/abell/.gemini/antigravity/brain/46a47252-b31e-41c5-8c48-2bca20c81f7e/cws_preview.png')
    print('CWS verified successfully!')

if __name__ == '__main__':
    test_cws()
