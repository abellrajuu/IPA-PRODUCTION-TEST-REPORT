import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import report_generator_python, pymupdf

def test_misc_table(cust_name):
    doc = docx.Document('DOC/Miscellaneous report.docx')
    t0 = doc.tables[0]
    c0 = t0.rows[0].cells[0]
    
    # Remove existing paragraphs in c0 except one or remove them from XML
    p_to_remove = list(c0.paragraphs)
    # Create a 2-column borderless table inside c0
    sub_table = c0.add_table(rows=3, cols=2)
    for p in p_to_remove:
        p._p.getparent().remove(p._p)
    # Remove borders from sub_table
    tblPr = sub_table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
    )
    tblPr.append(tblBorders)
    
    # Set cell widths: total width is ~540 pt (10800 dxa)
    # Col 0: 360 pt (7200 dxa), Col 1: 180 pt (3600 dxa)
    # Or Col 0: 320 pt (6400 dxa), Col 1: 220 pt (4400 dxa)
    col_w = [6400, 4400]
    for row in sub_table.rows:
        for c_i, w in enumerate(col_w):
            row.cells[c_i].width = int(w * 635)
            row.cells[c_i]._tc.tcPr.tcW.w = w
            
    # Populate cells
    data_rows = [
        (("Job No: ", "SY1615"), ("Model No: ", "SC-100")),
        (("Customer: ", cust_name), ("Sl. No: ", "11673")),
        (("Item Description: ", "Shear Pin Load Cell"), ("Qty: ", "1 No"))
    ]
    
    for r_idx, (col0_data, col1_data) in enumerate(data_rows):
        # Col 0
        p0 = sub_table.rows[r_idx].cells[0].paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        p0.paragraph_format.space_before = Pt(2)
        r_lbl0 = p0.add_run(col0_data[0])
        r_lbl0.bold = True
        r_lbl0.font.name = 'Times New Roman'
        r_lbl0.font.size = Pt(11)
        r_val0 = p0.add_run(col0_data[1])
        r_val0.bold = True
        r_val0.font.name = 'Times New Roman'
        r_val0.font.size = Pt(11)
        
        # Col 1
        p1 = sub_table.rows[r_idx].cells[1].paragraphs[0]
        p1.paragraph_format.space_after = Pt(2)
        p1.paragraph_format.space_before = Pt(2)
        r_lbl1 = p1.add_run(col1_data[0])
        r_lbl1.bold = True
        r_lbl1.font.name = 'Times New Roman'
        r_lbl1.font.size = Pt(11)
        r_val1 = p1.add_run(col1_data[1])
        r_val1.bold = True
        r_val1.font.name = 'Times New Roman'
        r_val1.font.size = Pt(11)

    out_docx = f'scratch/test_misc_{cust_name[:10]}.docx'
    out_pdf = f'scratch/test_misc_{cust_name[:10]}.pdf'
    doc.save(out_docx)
    report_generator_python.convert_docx_to_pdf(out_docx, out_pdf)
    doc_pdf = pymupdf.open(out_pdf)
    print(f'Cust: {cust_name[:20]} | Pages: {len(doc_pdf)}')
    for b in doc_pdf[0].get_text('dict')['blocks']:
        if b.get('type') == 0:
            for l in b['lines']:
                y = l['bbox'][1]
                if 50 < y < 140:
                    line_txt = ''.join(s['text'] for s in l['spans'])
                    print(f'  y={y:.1f}: {repr(line_txt)}')

test_misc_table('ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD')
