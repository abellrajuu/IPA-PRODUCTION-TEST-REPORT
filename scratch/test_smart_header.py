import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

def set_smart_header(p_job, p_cap, job_val, cust_val, cap_val, second_val, second_label="CONVEYOR NO:"):
    """
    Sets a perfectly aligned 2-column header where Customer Name NEVER wraps
    and is exactly vertically aligned with the second line (Conveyor No / Tag No).
    """
    # 1. Determine optimal tab position based on left column length
    left1 = f"JOB No: {job_val}"
    left2 = f"CAPACITY: {cap_val}"
    max_left_len = max(len(left1), len(left2))
    
    # 20 chars at 11pt is ~110pt = ~2200 dxa. Tab at 3800-4200 dxa gives ample room for both sides.
    tab_dxa = 3800 if len(cust_val) > 35 else (4200 if len(cust_val) > 25 else 4800)
    
    # 2. Determine font size for right column so it NEVER wraps
    # Usable width from tab_dxa to right margin (10811 dxa)
    avail_dxa = 10811 - tab_dxa
    avail_pt = avail_dxa / 20.0
    
    cust_full = f"CUSTOMER NAME: {cust_val}" if not cust_val.upper().startswith("CUSTOMER") else cust_val
    char_count = len(cust_full)
    
    # Average char width in Calibri/Arial at 11pt is ~5.6pt
    if char_count * 5.6 > avail_pt:
        font_size = max(8.0, round(avail_pt / (char_count * 0.55), 1))
    else:
        font_size = 11.0
        
    def build_para(p, left_text, right_text, right_font_size):
        p.text = ""
        pPr = p._p.get_or_add_pPr()
        for old_tabs in pPr.findall(qn('w:tabs')):
            pPr.remove(old_tabs)
        tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_dxa}"/></w:tabs>')
        pPr.append(tabs_xml)
        
        # Left run (11pt)
        r_left = p.add_run(f"{left_text}\t")
        r_left.font.size = Pt(11)
        r_left.bold = True
        
        # Right run (auto-scaled font)
        r_right = p.add_run(right_text)
        r_right.font.size = Pt(right_font_size)
        r_right.bold = True

    build_para(p_job, left1, cust_full, font_size)
    build_para(p_cap, left2, f"{second_label} {second_val}", 11.0)

# Test on a doc
d = docx.Document("DOC/Belt Scale.docx")
p3 = d.paragraphs[3]
p4 = d.paragraphs[4]

long_cust = "ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD"
set_smart_header(p3, p4, "JOB-SY1615", long_cust, "1200 TPH", "CV-101")
d.save("tests/test_smart_header.docx")
print("Saved tests/test_smart_header.docx successfully.")
