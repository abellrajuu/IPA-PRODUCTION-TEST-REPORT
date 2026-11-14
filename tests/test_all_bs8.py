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

base_data = {
    'report_type': 'belt_scale',
    'job_no': 'JOB-BS-8LC-TEST',
    'customer': 'Tata Steel Limited - Raw Material Division',
    'capacity': '1200 TPH',
    'conveyor_no': 'BC-201A',
    'belt_speed': '2.50 m/s',
    'bs_model': 'BS-600',
    'bs_serial': 'BS-2026-99',
    'sw_version': 'v4.1.0',
    'lc_model': 'LC-CZ-50T',
    'lc_system': '8',
    's1': 'LC-2026-001', 's2': 'LC-2026-002', 's3': 'LC-2026-003', 's4': 'LC-2026-004',
    's5': 'LC-2026-005', 's6': 'LC-2026-006', 's7': 'LC-2026-007', 's8': 'LC-2026-008',
    'ss_model': 'SS-100', 'ss_serial': 'SS-881',
    'as_model': 'AS-50', 'as_serial': 'AS-11',
    'jb_1_model': 'JB-8P', 'jb_1_serial': 'JB-001',
    'grid_data': [[f'{i*100}', f'{i*20}', '2.50', f'{i*100}', f'{i*50}', f'{4+i*1.6:.2f}', f'{4+i*1.6:.2f}', f'{4+i*1.6:.2f}'] for i in range(11)]
}

col_w = [2592, 3744, 2304, 2736]

for ao in ['1', '2', '3']:
    d_data = dict(base_data)
    d_data['num_outputs'] = ao
    test_docx = os.path.join(BASE_DIR, 'tests', f'test_bs8_{ao}ao.docx')
    test_pdf = os.path.join(BASE_DIR, 'tests', f'test_bs8_{ao}ao.pdf')
    rg.generate_docx_record(d_data, test_docx)
    
    # Apply 11pt and tab fix
    d = docx.Document(test_docx)
    t0 = d.tables[0]
    for r in t0.rows:
        for c_i, w in enumerate(col_w):
            r.cells[c_i].width = int(w * 635)
            r.cells[c_i]._tc.tcPr.tcW.w = w
    if hasattr(t0._tbl, 'tblGrid') and t0._tbl.tblGrid is not None:
        for c_i, w in enumerate(col_w):
            t0._tbl.tblGrid.gridCol_lst[c_i].w = w
            
    c1 = t0.rows[1].cells[1]
    pairs = [('LC-2026-001', 'LC-2026-005'), ('LC-2026-002', 'LC-2026-006'), ('LC-2026-003', 'LC-2026-007'), ('LC-2026-004', 'LC-2026-008')]
    for idx, (va, vb) in enumerate(pairs):
        p = c1.paragraphs[4 + idx]
        p.text = f"{idx+1}) {va}\t{idx+5}) {vb}"
        pPr = p._p.get_or_add_pPr()
        for old_tabs in pPr.findall(qn('w:tabs')):
            pPr.remove(old_tabs)
        tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1750"/></w:tabs>')
        pPr.append(tabs_xml)
        for r in p.runs:
            r.font.size = Pt(11)
            
    # Table 1 font 11pt
    for cp in d.tables[1].rows[0].cells[0].paragraphs:
        for r in cp.runs:
            r.font.size = Pt(11)
            
    d.save(test_docx)
    rg.convert_docx_to_pdf(test_docx, test_pdf)
    doc = fitz.open(test_pdf)
    print(f"Belt Scale 8LC {ao}AO page count: {len(doc)}")
