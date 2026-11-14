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

test_data = {
    'report_type': 'digital_indicator',
    'job_no': 'JOB-DI-2026-8LC',
    'customer': 'Tata Steel Limited',
    'capacity': '100 Tonnes',
    'tag_no': 'DI-TAG-01',
    'di_model': 'DI-3000',
    'di_serial': 'SN-DI-8821',
    'sw_version': 'v3.2.0',
    'lc_model': 'LC-CZ50T',
    'lc_system': '8',
    'calib_mode': 'load',
    'current_output': 'yes',
    's1': 'LC-SN-001', 's2': 'LC-SN-002', 's3': 'LC-SN-003', 's4': 'LC-SN-004',
    's5': 'LC-SN-005', 's6': 'LC-SN-006', 's7': 'LC-SN-007', 's8': 'LC-SN-008',
    'remote_model': 'N/A', 'remote_serial': 'N/A',
    'jbox_model': 'JB-8P', 'jbox_serial': 'JB-SN-441',
    'grid_data': [[f'{i*10}%'] + [f'{i*100}' for _ in range(8)] + [f'{4+i*1.6:.2f}'] for i in range(12)]
}

test_docx = os.path.join(BASE_DIR, 'tests', 'test_di8_11pt.docx')
test_pdf = os.path.join(BASE_DIR, 'tests', 'test_di8_11pt.pdf')

rg.generate_docx_record(test_data, test_docx)

d = docx.Document(test_docx)
t0 = d.tables[0]

# Col 0: 2500, Col 1: 3600 (SENSOR), Col 2: 2100 (REMOTE), Col 3: 3141 (JB)
col_w = [2500, 3600, 2100, 3141]
for r in t0.rows:
    for c_i, w in enumerate(col_w):
        r.cells[c_i].width = int(w * 635)
        r.cells[c_i]._tc.tcPr.tcW.w = w

if hasattr(t0._tbl, 'tblGrid') and t0._tbl.tblGrid is not None:
    for c_i, w in enumerate(col_w):
        t0._tbl.tblGrid.gridCol_lst[c_i].w = w

for r in t0.rows:
    for c in r.cells:
        for p in c.paragraphs:
            for run in p.runs:
                run.font.size = Pt(11)

c1 = t0.rows[1].cells[1]
pairs = [('LC-SN-001', 'LC-SN-005'), ('LC-SN-002', 'LC-SN-006'), ('LC-SN-003', 'LC-SN-007'), ('LC-SN-004', 'LC-SN-008')]
for idx, (va, vb) in enumerate(pairs):
    p = c1.paragraphs[4 + idx]
    p.text = f"{idx+1}) {va}\t{idx+5}) {vb}"
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')):
        pPr.remove(old_tabs)
    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1750"/></w:tabs>')
    pPr.append(tabs_xml)
    for run in p.runs:
        run.font.size = Pt(11)

for p in d.paragraphs:
    for run in p.runs:
        if run.font.size and run.font.size.pt in [9.5, 10.0, 10.5]:
            run.font.size = Pt(11)

for t in d.tables[1:]:
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for run in p.runs:
                    if run.font.size and run.font.size.pt in [9.5, 10.0]:
                        run.font.size = Pt(11)

d.save(test_docx)
rg.convert_docx_to_pdf(test_docx, test_pdf)

doc = fitz.open(test_pdf)
print("DI 8LC page count with 11pt:", len(doc))
for i in range(len(doc)):
    doc[i].get_pixmap(dpi=150).save(os.path.join(BASE_DIR, "tests", f"test_di8_11pt_p{i}.png"))
print("Saved test_di8_11pt PNGs successfully")
