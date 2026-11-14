import docx, os, sys
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, Inches
from report_generator_python import convert_docx_to_pdf
import pymupdf

def test_custom_table():
    doc = docx.Document('backend_docx_templates/PP-05_17 Misc - General Blank.docx')
    c1 = doc.tables[0].rows[1].cells[0]
    p0 = c1.paragraphs[0]
    p0.paragraph_format.space_after = Pt(6)

    headers = ['SL NO', 'ITEM / PARAMETER', 'SPECIFICATION', 'OBSERVED RESULT']
    num_rows = 5
    num_cols = len(headers)
    tbl = c1.add_table(rows=num_rows, cols=num_cols)
    tbl.style = 'Table Grid'

    total_width = Inches(6.8)
    widths = [Inches(0.8), Inches(2.8), Inches(1.8), Inches(1.4)]

    # Format headers
    for ci, h in enumerate(headers):
        c = tbl.rows[0].cells[ci]
        c.width = widths[ci]
        c.text = h
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E2E8F0"/>')
        c._tc.get_or_add_tcPr().append(shd)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.name = 'Arial'
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.bold = True

    for r_idx in range(1, num_rows):
        row_data = [str(r_idx), f'Custom Equipment Part #{r_idx}', '100% Operational', 'Passed / OK']
        for ci, val in enumerate(row_data):
            c = tbl.rows[r_idx].cells[ci]
            c.width = widths[ci]
            c.text = val
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER if ci in [0, 3] else docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
            if p.runs:
                p.runs[0].font.name = 'Arial'
                p.runs[0].font.size = Pt(8.5)

    p_end = c1.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.text = 'The above custom items were inspected, tested, and verified according to standard specifications and found OK.'
    p_end.runs[0].font.name = 'Arial'
    p_end.runs[0].font.size = Pt(9.5)

    to_delete = [p for p in c1.paragraphs if p._p != p0._p and p._p != p_end._p]
    for p in to_delete:
        p._p.getparent().remove(p._p)

    out_docx = 'tests/test_custom_table_styled.docx'
    doc.save(out_docx)
    convert_docx_to_pdf(out_docx)

    pdf = pymupdf.open('tests/test_custom_table_styled.pdf')
    print('Page count:', len(pdf))
    page = pdf[0]
    page.get_pixmap(dpi=150).save('C:/Users/abell/.gemini/antigravity/brain/46a47252-b31e-41c5-8c48-2bca20c81f7e/custom_table_styled_preview.png')
    print('Done!')

if __name__ == '__main__':
    test_custom_table()
