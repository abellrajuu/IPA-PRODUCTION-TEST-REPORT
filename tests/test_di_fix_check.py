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
    'job_no': 'JOB-DI-8LC-LOAD',
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
    'grid_data': [[f"{i*10}%"] + [f"{i*100}" for _ in range(8)] + [f"{4+i*1.6:.2f}"] for i in range(12)]
}

test_docx = os.path.join(BASE_DIR, 'tests', 'test_di_fix_check.docx')
test_pdf = os.path.join(BASE_DIR, 'tests', 'test_di_fix_check.pdf')

# Generate raw docx from template
d = docx.Document(os.path.join(BASE_DIR, 'backend_docx_templates', 'PP-05_06 Digital Indicator - 8 LC Load.docx'))

# Header
for p in d.paragraphs:
    txt = p.text
    if 'SYSTEM' in txt:
        p.text = '        SYSTEM: DIGITAL INDICATOR'
        for r in p.runs: r.bold = True; r.font.size = Pt(11)
    elif 'JOB No.' in txt:
        rg.set_para_with_tab(p, '        JOB No.: JOB-DI-8LC-LOAD', 'CUSTOMER NAME: Tata Steel Limited', 5200, 11, bold=True)
    elif 'CAPACITY:' in txt:
        rg.set_para_with_tab(p, '        CAPACITY: 100 Tonnes', 'TAG. NO: DI-TAG-01', 5200, 11, bold=True)
    elif 'Resolution:' in txt:
        p.text = '\tResolution: L: kg/T'
        for r in p.runs: r.bold = True; r.font.size = Pt(11)

# T0
t0 = d.tables[0]
col_w = [2500, 3600, 2100, 3141]
for r in t0.rows:
    for c_i, w in enumerate(col_w):
        r.cells[c_i].width = int(w * 635)
        r.cells[c_i]._tc.tcPr.tcW.w = w
for c_i, w in enumerate(col_w):
    t0._tbl.tblGrid.gridCol_lst[c_i].set(qn('w:w'), str(w))

for r in t0.rows:
    for c in r.cells:
        for p in c.paragraphs:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            for run in p.runs: run.font.size = Pt(11)

t0.cell(1, 0).text = 'Model No: DI-3000\n\nSerial No: SN-DI-8821\n\nS/W Version: v3.2.0'
for p in t0.cell(1, 0).paragraphs:
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
    for r in p.runs: r.font.size = Pt(11)

c1 = t0.cell(1, 1)
pairs = [('LC-SN-001', 'LC-SN-005'), ('LC-SN-002', 'LC-SN-006'), ('LC-SN-003', 'LC-SN-007'), ('LC-SN-004', 'LC-SN-008')]
for idx, (va, vb) in enumerate(pairs):
    p = c1.paragraphs[4 + idx]
    p.text = f"{idx+1}) {va}\t{idx+5}) {vb}"
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')): pPr.remove(old_tabs)
    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1750"/></w:tabs>')
    pPr.append(tabs_xml)
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
    for r in p.runs: r.font.size = Pt(11)

# T1
t1 = d.tables[1]
t1.cell(0, 0).text = 'Check for short between Line & Neutral, Neutral & Earth, Line & Earth: Checked & found OK'
for p in t1.cell(0, 0).paragraphs:
    for r in p.runs: r.font.size = Pt(9)

# T2
t2 = d.tables[2]
for r in t2.rows:
    for c in r.cells:
        for p in c.paragraphs:
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
            for r in p.runs: r.font.size = Pt(11)

# T3
t3 = d.tables[3]
grid_data = test_data['grid_data']
for r_idx in range(1, 13):
    row_cells = t3.rows[r_idx].cells
    row_arr = grid_data[r_idx - 1]
    row_cells[0].text = row_arr[0]
    for c_lc in range(1, 5): row_cells[c_lc].text = row_arr[c_lc]
    row_cells[5].text = row_arr[9]
    for c_extra in range(6, 10): row_cells[c_extra].text = row_arr[c_extra - 1]

for r in t3.rows:
    for c in r.cells:
        for p in c.paragraphs:
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.0
            for r in p.runs: r.font.size = Pt(11)

# T4, T5
for t in d.tables[4:]:
    for r in t.rows:
        for c in r.cells:
            for p in c.paragraphs:
                for r in p.runs: r.font.size = Pt(11)

# Remove empty paragraphs before routine test record
paras_to_remove = []
for i, p in enumerate(d.paragraphs):
    if 'ROUTINE TEST RECORD' in p.text and i > 0:
        p.paragraph_format.page_break_before = True
        for k in range(i - 1, -1, -1):
            if not d.paragraphs[k].text.strip():
                paras_to_remove.append(d.paragraphs[k])
        break

empty_before_inst = []
for i, p in enumerate(d.paragraphs):
    if 'Instrument Used' in p.text:
        for k in range(i - 1, -1, -1):
            if not d.paragraphs[k].text.strip():
                empty_before_inst.append(d.paragraphs[k])
                if len(empty_before_inst) >= 6: break
        break

for p in paras_to_remove + empty_before_inst:
    try: p._p.getparent().remove(p._p)
    except Exception: pass

d.save(test_docx)
rg.convert_docx_to_pdf(test_docx, test_pdf)
doc = fitz.open(test_pdf)
print('DI 8LC page count:', len(doc))
