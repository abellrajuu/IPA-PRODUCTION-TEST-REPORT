import os
import sys
import docx
from docx.shared import Pt, Inches
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
from werkzeug.utils import secure_filename
import threading
import copy
from datetime import datetime
import math
import re
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "backend_docx_templates")
EXPORTS_DIR = os.path.join(BASE_DIR, "backend_report_outputs")
SERVER_TEST_REPORT_DIR = r"\\192.168.100.248\prdndata\TEST REPORT"
os.makedirs(EXPORTS_DIR, exist_ok=True)
if not os.path.exists(SERVER_TEST_REPORT_DIR):
    try:
        os.makedirs(SERVER_TEST_REPORT_DIR, exist_ok=True)
    except Exception:
        pass

TEMPLATE_MAP = {
    "job_traveller_card": "PP-05_02 Job Traveller Card.docx",
    "belt_scale": "PP-05_03 Belt Scale.docx",
    "batching_system": "PP-05_04 Batching System.docx",
    "remote_indicator": "PP-05_05 Remote Indicator.docx",
    "digital_indicator": "PP-05_06 Digital Indicator - 4 LC Load.docx",
    "weigh_feeder": "PP-05_07 Weigh Feeder.docx",
    "crane_scale": "PP-05_08 Crane Scale.docx",
    "signal_conditioner": "PP-05_09 Signal Conditioner.docx",
    "trip_safe": "PP-05_10 tRIP sAFE.docx",
    "vibration_switch": "PP-05_11 Vibration Switch.docx",
    "acc_mv": "PP-05_12 ACC mV.docx",
    "acc_charge": "PP-05_13 ACC Charge.docx",
    "vibration_meter": "PP-05_14 Vibration Meter.docx",
    "charge_amplifier": "PP-05_15 Charge Amplifier.docx",
    "inprocess_register": "PP-05_16 Inprocess.docx",
    "inprocess": "PP-05_16 Inprocess.docx",
    "misc_report": "PP-05_17 Misc - Load Cell.docx",
    "misc": "PP-05_17 Misc - Load Cell.docx",
    "cws": "PP-05_08 Crane Scale.docx",
    "odd_system": "PP-05_18_A ODD System.docx",
    "dd_system": "PP-05_18_B DD System.docx",
    "work_instructions": "PP-05_19 Work Instructions.docx",
    "performance_index": "PP-05_20 Performance Index.docx",
    "performance_delay_analysis": "PP-05_21 Performance Delay Analysis.docx",
    "performance_delay": "PP-05_21 Performance Delay Analysis.docx",
    "equipment_list": "PP-05_22 Equipment List.docx",
    "loss_in_weigh_feeder": "PP-05_24 Loss in Weigh Feeder.docx",
}

def replace_in_paragraph(p, field_map):
    for key, val in field_map.items():
        placeholder = f"{{{{{key}}}}}"
        val_str = str(val) if val is not None else ""
        if placeholder in val_str:
            val_str = val_str.replace(placeholder, "")
        
        max_replacements = 50
        count = 0
        while count < max_replacements:
            count += 1
            t_elems = p._p.xpath('.//w:t')
            if not t_elems:
                break
                
            full_text = "".join(t.text for t in t_elems if t.text)
            if placeholder not in full_text:
                break
            
            char_mapping = []
            for t_idx, t in enumerate(t_elems):
                if t.text:
                    for char_idx in range(len(t.text)):
                        char_mapping.append((t_idx, char_idx))
                        
            start_idx = full_text.find(placeholder)
            end_idx = start_idx + len(placeholder) - 1
            
            # --- ALIGNMENT FIX: Space clearing after colon or sl number ---
            # When inserting val_str (e.g. 5 letters):
            # First clear 5 spaces following the placeholder / colon / sl number
            spaces_to_eat = len(val_str)
            while spaces_to_eat > 0 and end_idx + 1 < len(full_text) and full_text[end_idx + 1] in (' ', '_', '\xa0'):
                end_idx += 1
                spaces_to_eat -= 1

            # If still have spaces to eat and there are redundant leading spaces before placeholder
            # (e.g. after colon or sl number), eat leading spaces while preserving 1 space
            while spaces_to_eat > 0 and start_idx - 1 >= 0 and full_text[start_idx - 1] in (' ', '_', '\xa0'):
                if start_idx - 2 >= 0 and full_text[start_idx - 2] not in (':', ')', '.'):
                    start_idx -= 1
                    spaces_to_eat -= 1
                else:
                    break
            # --------------------------------------------------------------
            
            if start_idx < 0 or start_idx >= len(char_mapping) or end_idx >= len(char_mapping):
                break
            
            start_t_idx, start_char_idx = char_mapping[start_idx]
            end_t_idx, end_char_idx = char_mapping[end_idx]
            
            for i in range(start_t_idx + 1, end_t_idx):
                t_elems[i].text = ""
                
            if start_t_idx == end_t_idx:
                r_text = t_elems[start_t_idx].text
                t_elems[start_t_idx].text = r_text[:start_char_idx] + val_str + r_text[end_char_idx + 1:]
            else:
                r_start = t_elems[start_t_idx].text
                t_elems[start_t_idx].text = r_start[:start_char_idx] + val_str
                r_end = t_elems[end_t_idx].text
                t_elems[end_t_idx].text = r_end[end_char_idx + 1:]

def set_para_tab(p, text_left, text_right, tab_pos_dxa=4500, font_size_pt=9.5):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    p.text = f"{text_left}\t{text_right}"
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')):
        pPr.remove(old_tabs)
    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_pos_dxa}"/></w:tabs>')
    pPr.append(tabs_xml)
    for r in p.runs:
        r.bold = True
        r.font.size = Pt(font_size_pt)


import threading
com_lock = threading.Lock()

def convert_docx_to_pdf(docx_path, pdf_path=None):
    if not pdf_path:
        pdf_path = docx_path.replace(".docx", ".pdf")
    
    with com_lock:
        word = None
        doc = None
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            try:
                word.ActivePrinter = 'Microsoft Print to PDF'
            except Exception:
                pass
            
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 = wdFormatPDF
        except Exception as e:
            print(f"[PDF CONVERT WARNING] MS Word COM conversion failed for {docx_path}: {e}")
        finally:
            if doc:
                try: doc.Close(0)
                except Exception: pass
                try: del doc
                except Exception: pass
            if word:
                try: word.Quit()
                except Exception: pass
                try: del word
                except Exception: pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        # Secondary Fallback: Try docx2pdf ONLY after Word COM has been fully closed and uninitialized
        if not os.path.exists(pdf_path):
            try:
                from docx2pdf import convert
                convert(docx_path, pdf_path)
            except Exception as e2:
                print(f"[PDF CONVERT WARNING] Secondary fallback failed: {e2}")

    # Tertiary Fallback: Headless LibreOffice (for cross-platform and Linux/Word-less environments)
    if not os.path.exists(pdf_path):
        try:
            import subprocess
            import shutil
            soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
            if soffice_cmd:
                out_dir = os.path.dirname(os.path.abspath(pdf_path))
                subprocess.run([soffice_cmd, "--headless", "--convert-to", "pdf", "--outdir", out_dir, os.path.abspath(docx_path)], capture_output=True, timeout=45)
        except Exception as e3:
            print(f"[PDF CONVERT WARNING] LibreOffice headless fallback note: {e3}")

    if os.path.exists(pdf_path):
        try:
            if os.path.exists(SERVER_TEST_REPORT_DIR):
                import shutil
                shutil.copyfile(pdf_path, os.path.join(SERVER_TEST_REPORT_DIR, os.path.basename(pdf_path)))
        except Exception as e_cp:
            print(f"[NOTE] Could not mirror PDF to SERVER_TEST_REPORT_DIR: {e_cp}")
    else:
        print(f"[PDF CONVERT ERROR] PDF file was not created: {pdf_path}")
    return pdf_path

def clean_unhandled_tags(p):
    import re
    # Clear any remaining {{...}} tags to empty string
    max_clean = 50
    count = 0
    while count < max_clean:
        count += 1
        t_elems = p._p.xpath('.//w:t')
        if not t_elems:
            break
        full_text = "".join(t.text for t in t_elems if t.text)
        m = re.search(r'\{\{.*?\}\}', full_text)
        if not m:
            break
        placeholder = m.group(0)
        char_mapping = []
        for t_idx, t in enumerate(t_elems):
            if t.text:
                for char_idx in range(len(t.text)):
                    char_mapping.append((t_idx, char_idx))
        start_idx = full_text.find(placeholder)
        end_idx = start_idx + len(placeholder) - 1
        if start_idx < 0 or start_idx >= len(char_mapping) or end_idx >= len(char_mapping):
            break
        start_t_idx, start_char_idx = char_mapping[start_idx]
        end_t_idx, end_char_idx = char_mapping[end_idx]
        for i in range(start_t_idx + 1, end_t_idx):
            t_elems[i].text = ""
        if start_t_idx == end_t_idx:
            r_text = t_elems[start_t_idx].text
            t_elems[start_t_idx].text = r_text[:start_char_idx] + r_text[end_char_idx + 1:]
        else:
            r_start = t_elems[start_t_idx].text
            t_elems[start_t_idx].text = r_start[:start_char_idx]
            r_end = t_elems[end_t_idx].text
            t_elems[end_t_idx].text = r_end[end_char_idx + 1:]

def set_para_with_tab(p, text_left, text_right, tab_pos_dxa=4200, font_size_pt=11.0, bold=None):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    p.text = ""
    pPr = p._p.get_or_add_pPr()
    for old_tabs in pPr.findall(qn('w:tabs')):
        pPr.remove(old_tabs)
    for old_ind in pPr.findall(qn('w:ind')):
        pPr.remove(old_ind)

    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_pos_dxa}"/></w:tabs>')
    pPr.append(tabs_xml)

    # Hanging indent: ensures any wrapped text stays in the right column, NEVER wraps under text_left!
    ind_xml = parse_xml(f'<w:ind {nsdecls("w")} w:left="{tab_pos_dxa}" w:firstLine="-{tab_pos_dxa}"/>')
    pPr.append(ind_xml)

    # Left run
    r_left = p.add_run(f"{text_left}\t")
    if bold is not None:
        r_left.bold = bold
    r_left.font.size = Pt(font_size_pt)
    r_left.font.name = "Times New Roman"

    # Right run
    r_right = p.add_run(text_right)
    if bold is not None:
        r_right.bold = bold
    r_right.font.size = Pt(font_size_pt)
    r_right.font.name = "Times New Roman"

def _parse_first_float(val, default=None):
    """Extract the first valid decimal/signed float token cleanly without concatenating multiple dots."""
    if val is None:
        return default
    cleaned = re.sub(r'(?<=\d),(?=\d{3}\b)', '', str(val))
    m = re.search(r'-?\d+(?:\.\d+)?', cleaned)
    if not m:
        return default
    try:
        return float(m.group(0))
    except ValueError:
        return default


def _calc_ma_from_load(v0, default_x, span_val, unit_str, effective_sys_cap, is_load_calib):
    """Compute proportional 4-20mA current output handling unit conversions and % inputs safely."""
    cur_x = default_x
    if v0:
        v0_str = str(v0).strip()
        v0_n = _parse_first_float(v0_str)
        if v0_n is not None and v0_n >= 0:
            up_v0 = v0_str.upper()
            if '%' in up_v0:
                cur_x = (v0_n / 100.0) * (effective_sys_cap if is_load_calib and effective_sys_cap > 0 else span_val)
            elif is_load_calib:
                if 'KG' in up_v0 and unit_str == 'T':
                    cur_x = v0_n / 1000.0
                elif ('TON' in up_v0 or re.search(r'(?:^|\d|\s)T(?:\b|$)', up_v0)) and unit_str == 'Kg':
                    cur_x = v0_n * 1000.0
                else:
                    cur_x = v0_n
            else:
                cur_x = v0_n

    denom = effective_sys_cap if (is_load_calib and effective_sys_cap > 0) else (span_val if span_val > 0 else 10.0)
    ma_val = 4.0 + (16.0 / denom) * cur_x
    return f"{max(0.0, ma_val):.2f}"


def generate_digital_indicator_record(data, output_path):
    calib_mode = str(data.get("calib_mode", "mv")).lower().strip()
    raw_lc = str(data.get("lc_system") or "").strip().lower()
    has_8lc = any(str(data.get(f"s{i}") or data.get(f"lc_s{i}") or "").strip() for i in range(5, 9))
    m_lc = re.search(r'(\d+)', raw_lc)
    num_cells = int(m_lc.group(1)) if m_lc else (8 if has_8lc else 4)
    lc_system = "8" if (num_cells > 4 or has_8lc) else "4" 
    raw_ao = str(data.get("num_outputs") or data.get("num_ao") or "").strip()
    raw_cur = str(data.get("current_output") or "").lower().strip()
    has_ao2 = any(str(data.get(f"g{r}_5_ao2") or "").strip() for r in range(1, 13))

    if raw_cur in ["no", "none", "0"] or raw_ao == "0":
        num_ao = "0"
        cur_out_enabled = False
    elif raw_ao == "2" or raw_cur in ["2", "2 ao", "2 outputs"] or has_ao2:
        num_ao = "2"
        cur_out_enabled = True
    else:
        num_ao = "1"
        cur_out_enabled = True

    # Determine template
    if "load" in calib_mode:
        is_load_calib = True
        if num_ao == "2":
            tmpl_name = "PP-05_06 Digital Indicator - 8 LC Load - 2 AO.docx" if lc_system == "8" else "PP-05_06 Digital Indicator - 4 LC Load - 2 AO.docx"
        else:
            tmpl_name = "PP-05_06 Digital Indicator - 8 LC Load.docx" if lc_system == "8" else "PP-05_06 Digital Indicator - 4 LC Load.docx"
    else:
        is_load_calib = False
        if num_ao == "2":
            tmpl_name = "PP-05_06 Digital Indicator - 8 LC mV - 2 AO.docx" if lc_system == "8" else "PP-05_06 Digital Indicator - 4 LC mV - 2 AO.docx"
        else:
            tmpl_name = "PP-05_06 Digital Indicator - 8 LC mV.docx" if lc_system == "8" else "PP-05_06 Digital Indicator - 4 LC mV.docx"

    master_path = os.path.join(TEMPLATES_DIR, tmpl_name)
    if not os.path.exists(master_path):
        doc_candidate = os.path.join(BASE_DIR, "DOC", tmpl_name)
        if os.path.exists(doc_candidate):
            master_path = doc_candidate
        else:
            raise FileNotFoundError(f"Template not found: {master_path}")

    doc = docx.Document(master_path)

    def get_val(*keys, default=""):
        for k in keys:
            v = data.get(k)
            if v is not None and str(v).strip():
                return str(v).strip()
        return default

    job_no = get_val("job_no", "job", default="")
    customer = get_val("customer", "customer_name", "cust", default="")
    capacity = get_val("capacity", "cap", default="")
    tag_no = get_val("conveyor_no", "tag_no", "conveyor_tag", default="")
    resolution = get_val("resolution", "res_l", default="")

    di_model = get_val("di_model", "bs_model", "model_no", default="")
    di_serial = get_val("di_serial", "bs_serial", "serial_no", default="")
    sw_version = get_val("sw_version", "bs_sw", "version", default="")

    lc_model = get_val("lc_model", "sensor_model", default="")
    s1 = get_val("s1", default="")
    s2 = get_val("s2", default="")
    s3 = get_val("s3", default="")
    s4 = get_val("s4", default="")
    s5 = get_val("s5", default="")
    s6 = get_val("s6", default="")
    s7 = get_val("s7", default="")
    s8 = get_val("s8", default="")

    remote_model = get_val("remote_model", "rm_model", default="")
    remote_serial = get_val("remote_serial", "rm_serial", default="")

    jbox_model = get_val("jbox_model", "jbox_model1", "jb_1_model", default="")
    jbox_serial = get_val("jbox_serial", "jbox_serial1", "jb_1_serial", default="")
    jbox2_model = get_val("jbox2_model", "jbox_model2", "jb_2_model", default="")
    jbox2_serial = get_val("jbox2_serial", "jbox_serial2", "jb_2_serial", default="")

    short_check = get_val("short_check", "test_voltage_chk", default="")

    cust_str = str(customer or "")
    if len(cust_str) > 55:
        hdr_tab, hdr_pt = 3600, 9.5
    elif len(cust_str) > 40:
        hdr_tab, hdr_pt = 3900, 10.0
    else:
        hdr_tab, hdr_pt = 4200, 11.0

    # 1. Update Header Paragraphs
    for p in doc.paragraphs:
        txt = p.text
        if "SYSTEM  :" in txt or "SYSTEM:" in txt:
            p.text = "        SYSTEM  : DIGITAL INDICATOR"
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(11)
                r.font.name = "Times New Roman"
        elif "JOB No." in txt and ("CUSTOMER NAME:" in txt or "CUSTOMER:" in txt):
            set_para_with_tab(p, f"        JOB No.    : {job_no}", f"CUSTOMER NAME: {customer}", hdr_tab, hdr_pt, bold=True)
        elif "CAPACITY:" in txt and "TAG. NO" in txt:
            cap_str = capacity
            if cap_str and not any(u in cap_str.upper() for u in ["KG", "TON"]):
                cap_str = f"{cap_str} Kg"
            set_para_with_tab(p, f"        CAPACITY: {cap_str}", f"TAG. NO: {tag_no}", hdr_tab, hdr_pt, bold=True)
        elif "Resolution:" in txt:
            res_str = resolution if resolution else ""
            p.text = f"\tResolution: {res_str}" if res_str else "\tResolution: L: kg/T"
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(11)
        elif "REMARKS (If ANY):" in txt:
            remarks = get_val("remarks", default="")
            p.text = f"REMARKS (If ANY): - {remarks}"
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(11)
        elif "Instrument Used:" in txt:
            inst = get_val("instrument_used", default="")
            p.text = f"        Instrument Used: {inst}"
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                r.bold = True
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)
        elif "Tested by:" in txt and "Approved by:" in txt:
            tb = get_val("tested_by", default="")
            ab = get_val("approved_by", default="")
            set_para_with_tab(p, f"        Tested by: {tb}", f"Approved by: {ab}", 5400, 11, bold=True)
            p.paragraph_format.space_after = Pt(4)
        elif "Date:" in txt and "(HOD" in txt:
            dt = get_val("date", default=datetime.now().strftime("%d-%b-%Y"))
            set_para_with_tab(p, f"        Date: {dt}", "(HOD - PDN)", 5400, 11, bold=True)
            p.paragraph_format.space_after = Pt(8)

    # 2. Table 0 (Equipment Specs)
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        # Rebalance columns for 8 LC so SENSOR column has sufficient width for two 11pt serials
        if lc_system == "8":
            col_w = [2500, 3600, 2100, 3141]
            for r in t0.rows:
                if len(r.cells) >= 4:
                    for c_i, w in enumerate(col_w):
                        r.cells[c_i].width = int(w * 635)
                        r.cells[c_i]._tc.tcPr.tcW.w = w
            if hasattr(t0._tbl, 'tblGrid') and t0._tbl.tblGrid is not None and len(t0._tbl.tblGrid.gridCol_lst) >= 4:
                for c_i, w in enumerate(col_w):
                    t0._tbl.tblGrid.gridCol_lst[c_i].set(qn('w:w'), str(w))

        # Digital Indicator Cell
        t0.cell(1, 0).text = f"Model No: {di_model}\n\nSerial No: {di_serial}\n\nS/W Version: {sw_version}"

        # SENSOR Cell
        c1 = t0.cell(1, 1)
        if len(c1.paragraphs) >= 8:
            c1.paragraphs[1].text = f"Model No: {lc_model}" if lc_model else "Model No:"
            c1.paragraphs[3].text = "Serial No:"
            if lc_system == "8":
                pairs = [(s1, s5), (s2, s6), (s3, s7), (s4, s8)]
                for idx, (va, vb) in enumerate(pairs):
                    p_idx = 4 + idx
                    t_left = f"{idx+1}) {va}".strip() if va else f"{idx+1})"
                    t_right = f"{idx+5}) {vb}".strip() if vb else f"{idx+5})"
                    p_ser = c1.paragraphs[p_idx]
                    p_ser.text = f"{t_left}\t{t_right}"
                    pPr = p_ser._p.get_or_add_pPr()
                    for old_tabs in pPr.findall(qn('w:tabs')):
                        pPr.remove(old_tabs)
                    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1750"/></w:tabs>')
                    pPr.append(tabs_xml)
            else:
                singles = [s1, s2, s3, s4]
                for idx, va in enumerate(singles):
                    p_idx = 4 + idx
                    c1.paragraphs[p_idx].text = f"{idx+1}) {va}".strip() if va else f"{idx+1})"
        else:
            if lc_system == "8":
                c1.text = (
                    f"Model No: {lc_model}\n\nSerial No:\n"
                    f"1) {s1}\t5) {s5}\n"
                    f"2) {s2}\t6) {s6}\n"
                    f"3) {s3}\t7) {s7}\n"
                    f"4) {s4}\t8) {s8}"
                )
                for p_idx in range(4, min(8, len(c1.paragraphs))):
                    p_ser = c1.paragraphs[p_idx]
                    pPr = p_ser._p.get_or_add_pPr()
                    for old_tabs in pPr.findall(qn('w:tabs')):
                        pPr.remove(old_tabs)
                    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1750"/></w:tabs>')
                    pPr.append(tabs_xml)
            else:
                c1.text = (
                    f"Model No: {lc_model}\n\nSerial No:\n"
                    f"1) {s1}\n2) {s2}\n3) {s3}\n4) {s4}"
                )

        # REMOTE Cell
        rm_m = remote_model if remote_model else "N/A"
        rm_s = remote_serial if remote_serial else "N/A"
        t0.cell(1, 2).text = f"Model No: {rm_m}\n\nSerial No: {rm_s}"

        # JUNCTION BOX Cell (Supports dual junction boxes)
        jb1_m = jbox_model if jbox_model else "N/A"
        jb1_s = jbox_serial if jbox_serial else "N/A"
        jb2_m = jbox2_model if jbox2_model else ""
        jb2_s = jbox2_serial if jbox2_serial else ""

        c_jb = t0.cell(1, 3)
        if len(c_jb.paragraphs) >= 6:
            c_jb.paragraphs[1].text = f"Model No: {jb1_m}"
            c_jb.paragraphs[2].text = f"Serial No: {jb1_s}"
            c_jb.paragraphs[4].text = f"Model No: {jb2_m}" if jb2_m else ("Model No: N/A" if jb2_s else "Model No:")
            c_jb.paragraphs[5].text = f"Serial No: {jb2_s}" if jb2_s else ("Serial No: N/A" if jb2_m else "Serial No:")
        else:
            jb2_text = f"\n\nModel No: {jb2_m}\nSerial No: {jb2_s}" if (jb2_m or jb2_s) else ""
            c_jb.text = f"Model No: {jb1_m}\nSerial No: {jb1_s}{jb2_text}"

        # Apply 11pt and tight line spacing across Table 0
        for r in t0.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.0
                    for r_run in p.runs:
                        r_run.font.size = Pt(11)

    # 3. Table 1 (Short Check) - 9pt matching original master template
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        if short_check:
            short_display = "Checked & found OK" if ("ok" in short_check.lower() or "passed" in short_check.lower()) else short_check
        else:
            short_display = ""
        t1.cell(0, 0).text = f"Check for short between Line & Neutral, Neutral & Earth, Line & Earth: {short_display}"
        for cp in t1.cell(0, 0).paragraphs:
            for r in cp.runs:
                r.font.size = Pt(9)

    # 4. Table 2 (Routine Tests 1 to 5) - 11pt matching original template
    span_val = 10.0
    unit_str = "Kg" if is_load_calib else "mV"
    raw_spec4 = get_val("spec_4", "act_4", default="")
    raw_cap_fallback = get_val("capacity", "cap", default="")
    if is_load_calib and raw_spec4 and any(tok in raw_spec4.upper() for tok in ("MV", "NA", "OK")):
        raw_cap_str = raw_cap_fallback
    else:
        raw_cap_str = raw_spec4 or raw_cap_fallback
    if is_load_calib and raw_cap_str:
        num_cap = _parse_first_float(raw_cap_str)
        if num_cap is not None and num_cap > 0:
            span_val = num_cap
            up_cap = raw_cap_str.upper()
            if 'TON' in up_cap or ' T' in up_cap or up_cap.endswith('T'):
                unit_str = "T"
            else:
                unit_str = "Kg"
        elif not raw_cap_str:
            span_val = 1000.0
    elif not is_load_calib:
        num_cap = _parse_first_float(raw_cap_str)
        if num_cap is not None and 0 < num_cap <= 50 and 'MV' in raw_cap_str.upper():
            span_val = num_cap
        else:
            span_val = 10.0
        unit_str = "mV"

    dec = 2 if (unit_str == 'mV' or span_val < 100) else (1 if span_val < 1000 else 0)
    step_ratio = span_val / 5.0

    # Extract Total System Capacity (e.g. 268.0 T) for proportional current calculation
    raw_sys_cap = get_val("capacity", "cap", default="")
    sys_cap_num = 0.0
    sys_cap_unit = ""
    if raw_sys_cap:
        parsed_sys = _parse_first_float(raw_sys_cap, 0.0)
        if parsed_sys and parsed_sys > 0:
            sys_cap_num = parsed_sys
            up_sys = str(raw_sys_cap).upper()
            if 'TON' in up_sys or ' T' in up_sys or up_sys.endswith('T'): sys_cap_unit = "T"
            elif 'KG' in up_sys: sys_cap_unit = "Kg"

    effective_sys_cap = sys_cap_num if sys_cap_num > 0 else span_val
    if sys_cap_num > 0:
        if sys_cap_unit == 'T' and unit_str == 'Kg':
            effective_sys_cap = sys_cap_num * 1000.0
        elif sys_cap_unit == 'Kg' and unit_str == 'T':
            effective_sys_cap = sys_cap_num / 1000.0

    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for row_i in range(1, 6):
            if row_i < len(t2.rows):
                spec_v = get_val(f"spec_{row_i}", default="")
                act_v = get_val(f"act_{row_i}", default="")
                if not spec_v:
                    if row_i == 1: spec_v = "230V AC"
                    elif row_i == 2: spec_v = "24V DC"
                    elif row_i == 3: spec_v = "OK"
                    elif row_i == 4: spec_v = f"{span_val:.{dec}f} {unit_str}"
                    elif row_i == 5: spec_v = "NA"
                if not act_v:
                    if row_i == 1: act_v = "230V AC"
                    elif row_i == 2: act_v = "24V DC"
                    elif row_i == 3: act_v = "OK"
                    elif row_i == 4: act_v = f"{span_val:.{dec}f} {unit_str}"
                    elif row_i == 5: act_v = "NA"
                t2.cell(row_i, 2).text = spec_v
                t2.cell(row_i, 3).text = act_v
        for r in t2.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.0
                    for r in p.runs:
                        r.font.size = Pt(11)

    # 5. Table 3 (Calibration Matrix: 12 Rows) - 11pt with tight line spacing so Page 1 never overflows
    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        grid_data = data.get("grid_data")
        
        # Ensure header rows have tight spacing
        for hdr_r in t3.rows[:2 if num_ao == "2" else 1]:
            for cell in hdr_r.cells:
                for cp in cell.paragraphs:
                    cp.paragraph_format.space_before = Pt(0)
                    cp.paragraph_format.space_after = Pt(0)
                    cp.paragraph_format.line_spacing = 1.0
                    for r in cp.runs:
                        r.font.size = Pt(11)

        if num_ao == "2":
            # 2 AO template has 2 header rows, data starts at index 2 (rows 2 to 13)
            for r_idx in range(1, 13):
                r_target = r_idx + 1 # 2 to 13
                if r_target < len(t3.rows):
                    row_cells = t3.rows[r_target].cells
                    step_num = f"{(step_ratio * r_idx):.{dec}f}"
                    step_disp = f"{step_num} {unit_str}"
                    
                    # Col 0: Load
                    v0 = get_val(f"g{r_idx}_0")
                    if not v0 and isinstance(grid_data, list) and r_idx <= len(grid_data):
                        row_arr = grid_data[r_idx - 1]
                        if isinstance(row_arr, list) and len(row_arr) > 0:
                            v0 = str(row_arr[0] or "")
                    if not v0 and r_idx <= 5:
                        v0 = step_disp
                    row_cells[0].text = v0 if v0 else ""

                    # Current output mA: (16.0 / System Capacity) * Load + 4.00
                    default_x = (step_ratio * r_idx) if is_load_calib else ((r_idx / 5.0) * span_val)
                    ma_def = _calc_ma_from_load(v0, default_x, span_val, unit_str, effective_sys_cap, is_load_calib)
                    
                    # Col 1..4: LC1..LC4
                    for c_lc in range(1, 5):
                        v_lc = get_val(f"g{r_idx}_{c_lc}")
                        if not v_lc and isinstance(grid_data, list) and r_idx <= len(grid_data):
                            row_arr = grid_data[r_idx - 1]
                            if isinstance(row_arr, list) and len(row_arr) > c_lc:
                                v_lc = str(row_arr[c_lc] or "")
                        if not v_lc and r_idx <= 5:
                            v_lc = step_num
                        row_cells[c_lc].text = v_lc if v_lc else ""
                    
                    # Col 5: AO1
                    v_ao1 = get_val(f"g{r_idx}_5_ao1", f"g{r_idx}_5")
                    if not v_ao1 and isinstance(grid_data, list) and r_idx <= len(grid_data):
                        row_arr = grid_data[r_idx - 1]
                        if isinstance(row_arr, list) and len(row_arr) > 5:
                            v_ao1 = str(row_arr[5] or "")
                    if not v_ao1 and r_idx <= 5:
                        v_ao1 = ma_def
                    row_cells[5].text = v_ao1 if v_ao1 else ""
                    
                    # Col 6: AO2
                    v_ao2 = get_val(f"g{r_idx}_5_ao2")
                    if not v_ao2 and isinstance(grid_data, list) and r_idx <= len(grid_data):
                        row_arr = grid_data[r_idx - 1]
                        if isinstance(row_arr, list) and len(row_arr) > 6:
                            v_ao2 = str(row_arr[6] or "")
                    if not v_ao2 and r_idx <= 5:
                        v_ao2 = ma_def
                    row_cells[6].text = v_ao2 if v_ao2 else ""
                    
                    # Col 7..10: LC5..LC8 (if lc_system == '8')
                    if lc_system == "8":
                        for idx_lc8, c_extra in enumerate(range(7, 11), start=6):
                            v_extra = get_val(f"g{r_idx}_{c_extra}", f"g{r_idx}_{idx_lc8}")
                            if not v_extra and isinstance(grid_data, list) and r_idx <= len(grid_data):
                                row_arr = grid_data[r_idx - 1]
                                if isinstance(row_arr, list) and len(row_arr) > c_extra:
                                    v_extra = str(row_arr[c_extra] or "")
                            if not v_extra and r_idx <= 5:
                                v_extra = step_num
                            row_cells[c_extra].text = v_extra if v_extra else ""
                    else:
                        for c_extra in range(7, 11):
                            row_cells[c_extra].text = ""
                    
                    for col_idx in range(len(row_cells)):
                        for cp in row_cells[col_idx].paragraphs:
                            cp.paragraph_format.space_before = Pt(0)
                            cp.paragraph_format.space_after = Pt(0)
                            cp.paragraph_format.line_spacing = 1.0
                            for r in cp.runs:
                                r.font.size = Pt(11)
        else:
            # 1 AO template (or current output NO) - data starts at row 1 (rows 1 to 12)
            for r_idx in range(1, 13):
                if r_idx < len(t3.rows):
                    row_cells = t3.rows[r_idx].cells
                    step_num = f"{(step_ratio * r_idx):.{dec}f}"
                    step_disp = f"{step_num} {unit_str}"
                    
                    # Col 0: Load
                    v0 = get_val(f"g{r_idx}_0")
                    if not v0 and isinstance(grid_data, list) and r_idx <= len(grid_data):
                        row_arr = grid_data[r_idx - 1]
                        if isinstance(row_arr, list) and len(row_arr) > 0:
                            v0 = str(row_arr[0] or "")
                    if not v0 and r_idx <= 5:
                        v0 = step_disp
                    row_cells[0].text = v0 if v0 else ""

                    # Current output mA: (16.0 / System Capacity) * Load + 4.00
                    default_x = (step_ratio * r_idx) if is_load_calib else ((r_idx / 5.0) * span_val)
                    ma_def = _calc_ma_from_load(v0, default_x, span_val, unit_str, effective_sys_cap, is_load_calib)
                    
                    # Col 1..4: LC1..LC4
                    for c_lc in range(1, 5):
                        v_lc = get_val(f"g{r_idx}_{c_lc}")
                        if not v_lc and isinstance(grid_data, list) and r_idx <= len(grid_data):
                            row_arr = grid_data[r_idx - 1]
                            if isinstance(row_arr, list) and len(row_arr) > c_lc:
                                v_lc = str(row_arr[c_lc] or "")
                        if not v_lc and r_idx <= 5:
                            v_lc = step_num
                        row_cells[c_lc].text = v_lc if v_lc else ""
                    
                    # Col 5: Current output mA (blank if current_output is 'no')
                    if cur_out_enabled:
                        v_ma = get_val(f"g{r_idx}_5_ao1", f"g{r_idx}_5")
                        if not v_ma and isinstance(grid_data, list) and r_idx <= len(grid_data):
                            row_arr = grid_data[r_idx - 1]
                            if isinstance(row_arr, list) and len(row_arr) > 5:
                                v_ma = str(row_arr[5] or "")
                        if not v_ma and r_idx <= 5:
                            v_ma = ma_def
                        row_cells[5].text = v_ma if v_ma else ""
                    else:
                        row_cells[5].text = ""
                    
                    # Col 6..9: LC5..LC8 (if lc_system == '8')
                    if lc_system == "8":
                        for c_extra in range(6, 10):
                            v_extra = get_val(f"g{r_idx}_{c_extra}")
                            if not v_extra and isinstance(grid_data, list) and r_idx <= len(grid_data):
                                row_arr = grid_data[r_idx - 1]
                                if isinstance(row_arr, list) and len(row_arr) > c_extra:
                                    v_extra = str(row_arr[c_extra] or "")
                            if not v_extra and r_idx <= 5:
                                v_extra = step_num
                            row_cells[c_extra].text = v_extra if v_extra else ""
                    else:
                        for c_extra in range(6, 10):
                            row_cells[c_extra].text = ""
                    
                    for col_idx in range(len(row_cells)):
                        for cp in row_cells[col_idx].paragraphs:
                            cp.paragraph_format.space_before = Pt(0)
                            cp.paragraph_format.space_after = Pt(0)
                            cp.paragraph_format.line_spacing = 1.0
                            for r in cp.runs:
                                r.font.size = Pt(11)

    # 6. Table 4 (Checklist 6 to 10) - 11pt
    if len(doc.tables) > 4:
        t4 = doc.tables[4]
        for row_i in range(1, 6):
            if row_i < len(t4.rows):
                if row_i == 5:
                    # Item 10: Rating of the fuse used in the system
                    # Next to 'Rating of the fuse used in the system:' put the rating (e.g. 1A)
                    # On the right side (Checked: YES/NO column) put 'YES'
                    fuse_rating = get_val("spec_10", default="")
                    if not fuse_rating or fuse_rating.upper() in ("OK", "YES", "CHECKED", "PASSED"):
                        raw_c10 = get_val("check_10", "act_10", default="")
                        if any(c in raw_c10.upper() for c in ("A", "MA")):
                            fuse_rating = raw_c10
                        else:
                            fuse_rating = "1A"
                    
                    t4.cell(row_i, 1).text = f"Rating of the fuse used in the system: {fuse_rating}"
                    for cp in t4.cell(row_i, 1).paragraphs:
                        cp.paragraph_format.space_before = Pt(0)
                        cp.paragraph_format.space_after = Pt(0)
                        cp.paragraph_format.line_spacing = 1.0
                        for r in cp.runs:
                            r.font.name = "Times New Roman"
                            r.font.size = Pt(11)

                    fuse_chk = get_val("check_10", "act_10", default="YES")
                    if any(c in fuse_chk.upper() for c in ("A", "MA")):
                        fuse_chk = "YES"
                    t4.cell(row_i, 2).text = fuse_chk
                    for cp in t4.cell(row_i, 2).paragraphs:
                        cp.paragraph_format.space_before = Pt(0)
                        cp.paragraph_format.space_after = Pt(0)
                        cp.paragraph_format.line_spacing = 1.0
                        for r in cp.runs:
                            r.font.name = "Times New Roman"
                            r.font.size = Pt(11)
                else:
                    default_chk = "OK" if row_i in (1, 2, 3) else "NA"
                    chk_v = get_val(f"check_{row_i + 5}", f"act_{row_i + 5}", f"spec_{row_i + 5}", default=default_chk)
                    if chk_v:
                        t4.cell(row_i, 2).text = chk_v
                    for cp in t4.cell(row_i, 2).paragraphs:
                        cp.paragraph_format.space_before = Pt(0)
                        cp.paragraph_format.space_after = Pt(0)
                        cp.paragraph_format.line_spacing = 1.0
                        for r in cp.runs:
                            r.font.name = "Times New Roman"
                            r.font.size = Pt(11)

    # 7. Table 5 (SPAN, TARE, EXC V - Load Cell mode only) - 11pt, clean centered alignment
    if is_load_calib and len(doc.tables) > 5:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        t5 = doc.tables[5]
        max_lc = 8 if lc_system == "8" else 4

        # Clean Row 0 Header: LC 1 to LC 4 (or LC 8) without leading empty lines
        t5.cell(0, 0).text = ""
        for lc_idx in range(1, max_lc + 1):
            if lc_idx < len(t5.columns):
                t5.cell(0, lc_idx).text = f"LC {lc_idx}"
                for cp in t5.cell(0, lc_idx).paragraphs:
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cp.paragraph_format.space_before = Pt(0)
                    cp.paragraph_format.space_after = Pt(0)
                    cp.paragraph_format.line_spacing = 1.0
                    for r in cp.runs:
                        r.bold = True
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(11)

        # Row 1: SPAN
        t5.cell(1, 0).text = "SPAN"
        for cp in t5.cell(1, 0).paragraphs:
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(0)
            cp.paragraph_format.space_after = Pt(0)
            cp.paragraph_format.line_spacing = 1.0
            for r in cp.runs:
                r.bold = True
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        for lc_idx in range(1, max_lc + 1):
            if lc_idx < len(t5.columns):
                val_span = get_val(f"span_lc{lc_idx}", f"span_{lc_idx}", default="")
                if not val_span:
                    val_span = f"{span_val:.{dec}f}"
                t5.cell(1, lc_idx).text = val_span
                for cp in t5.cell(1, lc_idx).paragraphs:
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cp.paragraph_format.space_before = Pt(0)
                    cp.paragraph_format.space_after = Pt(0)
                    cp.paragraph_format.line_spacing = 1.0
                    for r in cp.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(11)

        # Row 2: TARE
        t5.cell(2, 0).text = "TARE"
        for cp in t5.cell(2, 0).paragraphs:
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(0)
            cp.paragraph_format.space_after = Pt(0)
            cp.paragraph_format.line_spacing = 1.0
            for r in cp.runs:
                r.bold = True
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        for lc_idx in range(1, max_lc + 1):
            if lc_idx < len(t5.columns):
                val_tare = get_val(f"tare_lc{lc_idx}", f"tare_{lc_idx}", default="0.00")
                t5.cell(2, lc_idx).text = val_tare
                for cp in t5.cell(2, lc_idx).paragraphs:
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cp.paragraph_format.space_before = Pt(0)
                    cp.paragraph_format.space_after = Pt(0)
                    cp.paragraph_format.line_spacing = 1.0
                    for r in cp.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(11)

        # Row 3: EXC V
        t5.cell(3, 0).text = "EXC V"
        for cp in t5.cell(3, 0).paragraphs:
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(0)
            cp.paragraph_format.space_after = Pt(0)
            cp.paragraph_format.line_spacing = 1.0
            for r in cp.runs:
                r.bold = True
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        for lc_idx in range(1, max_lc + 1):
            if lc_idx < len(t5.columns):
                val_excv = get_val(f"excv_lc{lc_idx}", f"excv_{lc_idx}", default="10.0V")
                t5.cell(3, lc_idx).text = val_excv
                for cp in t5.cell(3, lc_idx).paragraphs:
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cp.paragraph_format.space_before = Pt(0)
                    cp.paragraph_format.space_after = Pt(0)
                    cp.paragraph_format.line_spacing = 1.0
                    for r in cp.runs:
                        r.font.name = "Times New Roman"
                        r.font.size = Pt(11)

    # 8. Clean up excess empty spacer paragraphs and enforce clean page transition
    paras_to_remove = []
    # Remove empty spacer paragraphs between Table 3 and Table 4, and lock heading to Page 2
    for i, p in enumerate(doc.paragraphs):
        if "ROUTINE TEST RECORD" in p.text and i > 0:
            p.paragraph_format.page_break_before = True
            for k in range(i - 1, 0, -1):
                if not doc.paragraphs[k].text.strip():
                    paras_to_remove.append(doc.paragraphs[k])
                else:
                    break
            break

    # Remove extra empty spacer paragraphs before Instrument Used (between Table 5 and footer)
    empty_before_inst = []
    for i, p in enumerate(doc.paragraphs):
        if "Instrument Used" in p.text:
            for k in range(i - 1, 0, -1):
                if not doc.paragraphs[k].text.strip():
                    empty_before_inst.append(doc.paragraphs[k])
                    if len(empty_before_inst) >= 6:
                        break
                else:
                    break
            break

    for p in paras_to_remove + empty_before_inst:
        try:
            p._p.getparent().remove(p._p)
        except Exception:
            pass

    # Save to output path
    doc.save(output_path)

    # Convert to PDF
    convert_docx_to_pdf(output_path)
    return output_path

def generate_signal_conditioner_record(data, output_path):
    calib_type = str(data.get("calib_type") or data.get("calib_mode") or "mv").lower().strip()
    is_lc_calib = "load" in calib_type

    if is_lc_calib:
        tmpl_name = "PP-05_09 Signal Conditioner - Load Cell.docx"
    else:
        tmpl_name = "PP-05_09 Signal Conditioner.docx"

    master_path = os.path.join(TEMPLATES_DIR, tmpl_name)
    if not os.path.exists(master_path):
        doc_candidate = os.path.join(BASE_DIR, "DOC", tmpl_name)
        if os.path.exists(doc_candidate):
            master_path = doc_candidate
        else:
            raise FileNotFoundError(f"Template not found: {master_path}")

    doc = docx.Document(master_path)

    def get_val(*keys, default=""):
        for k in keys:
            v = data.get(k)
            if v is not None and str(v).strip():
                return str(v).strip()
        return default

    job_no = get_val("job_no", "job", default="")
    customer = get_val("customer", "customer_name", "cust", default="")
    sc_model = get_val("sc_model", "model_no", default="")
    sc_serial = get_val("sc_serial", "serial_no", default="")
    lc_model = get_val("lc_model", default="")
    lc_serial = get_val("lc_serial", default="")

    span = get_val("span", default="")
    tare = get_val("tare", default="")
    exc_v = get_val("exc_v", "excitation_voltage", default="")

    annunciation = get_val("annunciation", default="")
    pct_error = get_val("pct_error", "error_percent", default="")
    multimeter_used = get_val("multimeter_used", "multimeter", default="")
    tested_by = get_val("tested_by", "user_name", default="")
    approved_by = get_val("approved_by", default="")
    date_val = get_val("date", default="")
    if not date_val:
        date_val = datetime.now().strftime("%d-%b-%Y")

    def find_p(predicate):
        for p in doc.paragraphs:
            if predicate(p.text):
                return p
        return None

    def set_para_2col(p, label1, val1, label2, val2, tab1_pos=850, tab2_pos=5102, font_name='Times New Roman', font_size=11, **kwargs):
        if not p: return
        p.text = ''
        pPr = p._p.get_or_add_pPr()
        existing_tabs = pPr.find(qn('w:tabs'))
        if existing_tabs is not None:
            pPr.remove(existing_tabs)
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
        for pos in [tab1_pos, tab2_pos]:
            tab = OxmlElement('w:tab')
            tab.set(qn('w:val'), 'left')
            tab.set(qn('w:pos'), str(pos))
            tabs.append(tab)

        lbl1_str = label1.strip()
        if not lbl1_str.endswith(':'): lbl1_str += ':'
        r1 = p.add_run('\t' + lbl1_str + ' ')
        r1.bold = True
        r1.font.name = font_name
        r1.font.size = Pt(font_size)

        r_v1 = p.add_run(str(val1).strip() if val1 else '')
        r_v1.bold = False
        r_v1.font.name = font_name
        r_v1.font.size = Pt(font_size)

        lbl2_str = label2.strip()
        if not lbl2_str.endswith(':'): lbl2_str += ':'
        r2 = p.add_run('\t' + lbl2_str + ' ')
        r2.bold = True
        r2.font.name = font_name
        r2.font.size = Pt(font_size)

        v2_str = str(val2).strip() if val2 else ''
        r_v2 = p.add_run(v2_str)
        r_v2.bold = False
        r_v2.font.name = font_name
        r_v2.font.size = Pt(font_size)

    def set_para_1col(p, label, val, tab1_pos=850, font_name='Times New Roman', font_size=11, **kwargs):
        if not p: return
        p.text = ''
        pPr = p._p.get_or_add_pPr()
        existing_tabs = pPr.find(qn('w:tabs'))
        if existing_tabs is not None:
            pPr.remove(existing_tabs)
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'left')
        tab.set(qn('w:pos'), str(tab1_pos))
        tabs.append(tab)

        lbl_str = label.strip()
        if not lbl_str.endswith(':'): lbl_str += ':'
        r = p.add_run('\t' + lbl_str + ' ')
        r.bold = True
        r.font.name = font_name
        r.font.size = Pt(font_size)

        r_v = p.add_run(str(val).strip() if val else '')
        r_v.bold = False
        r_v.font.name = font_name
        r_v.font.size = Pt(font_size)

    # Populate Metadata
    set_para_1col(find_p(lambda t: "job order no" in t.lower()), 'JOB ORDER No:', job_no)
    set_para_1col(find_p(lambda t: "customer" in t.lower() and "customer name" not in t.lower() or "customer:" in t.lower() or "customer name:" in t.lower()), 'CUSTOMER:', customer)

    p_sc_lc = find_p(lambda t: "signal conditioner" in t.lower() and "load cell" in t.lower() and "test record" not in t.lower())
    if p_sc_lc:
        p_sc_lc.text = ''
        pPr = p_sc_lc._p.get_or_add_pPr()
        existing_tabs = pPr.find(qn('w:tabs'))
        if existing_tabs is not None:
            pPr.remove(existing_tabs)
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
        for pos in [850, 5102]:
            tab = OxmlElement('w:tab')
            tab.set(qn('w:val'), 'left')
            tab.set(qn('w:pos'), str(pos))
            tabs.append(tab)
        r1 = p_sc_lc.add_run('\tSIGNAL CONDITIONER\tLOAD CELL')
        r1.bold = True
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(11)

    set_para_2col(find_p(lambda t: "model no:" in t.lower()), 'MODEL NO:', sc_model, 'MODEL NO:', lc_model)
    set_para_2col(find_p(lambda t: "serial no:" in t.lower()), 'SERIAL NO:', sc_serial, 'SERIAL NO:', lc_serial)

    # Populate Linearity Table (Table 0) - 5 equal step divisions (6 points: 0%, 20%, 40%, 60%, 80%, 100%)
    span_str = str(span or "").strip()
    span_num = _parse_first_float(span_str) if span_str else None
    if span_num is None or span_num <= 0:
        span_num = 20.0 if not is_lc_calib else 1000.0

    in_unit = " mV" if not is_lc_calib else " kg"
    out_unit = " mA"

    if doc.tables:
        t = doc.tables[0]
        for i in range(1, 7):
            if i < len(t.rows):
                in_val = get_val(f"cal_in_{i}", default="")
                out_val = get_val(f"cal_out_{i}", default="")
                if not in_val and isinstance(data.get("grid_data"), list) and i-1 < len(data["grid_data"]):
                    row_data = data["grid_data"][i-1]
                    if len(row_data) >= 2:
                        in_val = str(row_data[0])
                        out_val = str(row_data[1])

                # Default 5-step calculations if values are empty
                if not in_val:
                    pct = (i - 1) * 0.20
                    in_val = f"{span_num * pct:.2f}{in_unit}"
                if not out_val:
                    pct = (i - 1) * 0.20
                    out_val = f"{4.00 + 16.00 * pct:.2f}{out_unit}"

                row_cells = t.rows[i].cells
                row_cells[0].text = ''
                p0 = row_cells[0].paragraphs[0]
                p0.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                r_sl = p0.add_run(str(i))
                r_sl.font.name = 'Times New Roman'
                r_sl.font.size = Pt(11)

                row_cells[1].text = ''
                p1 = row_cells[1].paragraphs[0]
                p1.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                r_in = p1.add_run(str(in_val))
                r_in.font.name = 'Times New Roman'
                r_in.font.size = Pt(11)

                row_cells[2].text = ''
                p2 = row_cells[2].paragraphs[0]
                p2.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                r_out = p2.add_run(str(out_val))
                r_out.font.name = 'Times New Roman'
                r_out.font.size = Pt(11)

    # If Load Cell Calibration, populate Table 1 (SPAN, TARE, EXC V)
    if len(doc.tables) > 1 and is_lc_calib:
        t1 = doc.tables[1]
        def set_c(cell, val_text):
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(val_text))
            r.font.name = 'Times New Roman'
            r.font.size = Pt(11)

        if len(t1.rows) > 1:
            if len(t1.rows[1].cells) >= 3:
                set_c(t1.rows[1].cells[0], span or f"{span_num:.2f}{in_unit}")
                set_c(t1.rows[1].cells[1], tare or f"0.00{in_unit}")
                set_c(t1.rows[1].cells[2], exc_v or "10.0 VDC")

    # Functional Checks & Multimeter
    set_para_1col(find_p(lambda t: "power on annunciation" in t.lower()), 'Check the POWER ON Annunciation:', annunciation or "OK", tab1_pos=850, val_pos=4200)
    set_para_1col(find_p(lambda t: "percentage of error" in t.lower()), 'Percentage of error:', pct_error or "< 0.05%", tab1_pos=850, val_pos=3000)
    set_para_1col(find_p(lambda t: "multimeter used" in t.lower()), 'Multimeter Used:', multimeter_used or "Fluke 179 True RMS Multimeter", tab1_pos=850, val_pos=3000)

    # Sign-Off
    set_para_2col(find_p(lambda t: "tested by:" in t.lower()), 'Tested by:', tested_by, 'Approved by:', approved_by, tab1_pos=850, val1_pos=2200, tab2_pos=5102, val2_pos=6400)
    set_para_2col(find_p(lambda t: "date:" in t.lower()), 'Date:', date_val, '(HOD - PDN)', '', tab1_pos=850, val1_pos=2200, tab2_pos=5102, val2_pos=6400)

    doc.save(output_path)
    convert_docx_to_pdf(output_path)
    return output_path

def generate_misc_record(data, output_path):
    def get_val(*keys, default=""):
        for k in keys:
            v = data.get(k)
            if v is not None and str(v).strip():
                return str(v).strip()
        return default

    # 1. Check subsystem flags & explicit category
    category = str(data.get("misc_type") or data.get("misc_category") or "").lower().strip()
    chk_lc = str(data.get("chk_loadcell") or "").lower().strip() in ["1", "true", "on", "yes"]
    chk_jb = str(data.get("chk_jbox") or "").lower().strip() in ["1", "true", "on", "yes"]
    chk_rd = str(data.get("chk_remote") or "").lower().strip() in ["1", "true", "on", "yes"]
    chk_ind = str(data.get("chk_indicator") or "").lower().strip() in ["1", "true", "on", "yes"]
    chk_custom = str(data.get("chk_custom") or "").lower().strip() in ["1", "true", "on", "yes"]

    custom_table_raw = data.get("custom_table_json") or data.get("custom_table_data") or data.get("custom_table") or None
    custom_table = None
    if custom_table_raw:
        try:
            if isinstance(custom_table_raw, str):
                custom_table = json.loads(custom_table_raw)
            elif isinstance(custom_table_raw, dict):
                custom_table = custom_table_raw
        except Exception:
            custom_table = None

    num_checked = sum([chk_lc, chk_jb, chk_rd, chk_ind, chk_custom])

    item_desc_lower = str(get_val("custom_desc", "item_desc", "item_description", "description", default="")).lower()
    model_no_lower = str(get_val("custom_model", "model_no", "model", default="")).lower()

    # Intelligent Mode Resolution
    if chk_custom or (custom_table and custom_table.get("rows")) or "custom" in category:
        mode = "general"
    elif num_checked > 1:
        mode = "general"
    elif "load" in category or (not category and chk_lc) or any(data.get(f"lc_serial_{i}") or data.get(f"lc_sl_{i}") for i in range(1, 9)) or "load cell" in item_desc_lower or "load cell" in model_no_lower:
        mode = "load_cell"
    elif "remote" in category or "display" in category or (not category and chk_rd) or data.get("remote_model") or "remote" in item_desc_lower or "remote" in model_no_lower:
        mode = "remote_display"
    elif "multiple" in category or "multi" in category:
        mode = "jb_multiple"
    elif "single" in category:
        mode = "jb_single"
    elif "junction" in category or "jb" in category or (not category and chk_jb) or data.get("jbox_model_1") or "junction" in item_desc_lower or "jbox" in item_desc_lower:
        has_multi = bool(data.get("variant_model_2") or data.get("jbox_model_2") or data.get("variant_serial_2") or data.get("jbox_serials_2"))
        mode = "jb_multiple" if has_multi else "jb_single"
    elif "general" in category or "indicator" in category or (not category and chk_ind) or data.get("ind_model"):
        mode = "general"
    else:
        mode = "general"

    if mode == "load_cell":
        tmpl_name = "PP-05_17 Misc - Load Cell.docx"
    elif mode == "remote_display":
        tmpl_name = "PP-05_17 Misc - Remote Display.docx"
    elif mode == "jb_multiple":
        tmpl_name = "PP-05_17 Misc - Junction Box Multiple.docx"
    elif mode == "jb_single":
        tmpl_name = "PP-05_17 Misc - Junction Box Single.docx"
    else:
        tmpl_name = "PP-05_17 Misc - General Blank.docx"
        mode = "general"

    master_path = os.path.join(TEMPLATES_DIR, tmpl_name)
    if not os.path.exists(master_path):
        doc_candidate = os.path.join(BASE_DIR, "DOC", tmpl_name)
        if os.path.exists(doc_candidate):
            master_path = doc_candidate
        else:
            raise FileNotFoundError(f"Template not found: {master_path}")

    doc = docx.Document(master_path)

    def set_para_tab(p, left_text, right_text, tab_pos=5400, font_size=11, bold_label=True):
        if not p: return
        p.text = ""
        pPr = p._p.get_or_add_pPr()
        for old_tabs in pPr.findall(qn('w:tabs')):
            pPr.remove(old_tabs)
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="{tab_pos}"/></w:tabs>')
        pPr.append(tabs_xml)

        r1 = p.add_run(f"{left_text}\t")
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(font_size)
        r1.bold = bold_label

        r2 = p.add_run(str(right_text))
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(font_size)
        r2.bold = bold_label

    job_no = get_val("job_no", "job", default="")
    customer = get_val("customer", "customer_name", "cust", default="")
    item_desc = get_val("custom_desc", "item_desc", "item_description", "description", default="")
    qty = get_val("custom_qty", "qty", "quantity", default="")
    model_no = get_val("custom_model", "model_no", "model", default="")

    # Parse serial numbers (support custom_sl, sl_no, serials, serial_no, serial)
    raw_serials = data.get("custom_sl") or data.get("serials") or data.get("serial_no") or data.get("serial") or data.get("sl_no") or ""
    if isinstance(raw_serials, list):
        serials_list = [str(s).strip() for s in raw_serials if str(s).strip()]
    else:
        serials_list = [s.strip() for s in str(raw_serials).replace('\n', ',').split(',') if s.strip()]

    # If serials_list is empty, gather from subsystem specific fields
    if not serials_list:
        if mode == "load_cell":
            serials_list = [str(get_val(f"lc_serial_{i}", f"lc_sl_{i}")).strip() for i in range(1, 9) if get_val(f"lc_serial_{i}", f"lc_sl_{i}").strip()]
        elif mode in ["jb_single", "jb_multiple"]:
            for i in range(1, 6):
                s_str = get_val(f"jbox_serials_{i}", f"variant_serial_{i}")
                if s_str:
                    serials_list.extend([s.strip() for s in s_str.split(',') if s.strip()])
        elif mode == "remote_display":
            s_str = get_val("remote_serial", "remote_serial_input")
            if s_str:
                serials_list.append(s_str)

    serial_count = len(serials_list)
    serials_str = ", ".join(serials_list)

    if not qty and serial_count > 0:
        qty = f"{serial_count} Nos"

    # Default description if blank
    if not item_desc:
        if mode == "load_cell": item_desc = "Load Cell Subsystem"
        elif mode == "remote_display": item_desc = "Remote Display Unit"
        elif mode in ["jb_single", "jb_multiple"]: item_desc = "Junction Box Subsystem"
        elif mode == "general": item_desc = "Miscellaneous Subsystem(s)"

    # Default model if blank
    if not model_no:
        if mode == "remote_display": model_no = get_val("remote_model", "remote_model_input")
        elif mode == "load_cell": model_no = "Load Cell (3 Boxes Test)"
        elif mode in ["jb_single", "jb_multiple"]: model_no = get_val("jbox_model_1", "variant_model_1", "jbox_model")

    # 1. Header Box (Table 0 Row 0 Cell 0) - 2-column borderless table prevents line-wrap distortion
    c0 = doc.tables[0].rows[0].cells[0]
    if mode == "jb_multiple":
        hdr_model = ""
        hdr_serial = ""
    elif mode in ["remote_display", "jb_single"]:
        hdr_model = model_no
        hdr_serial = serials_str if serial_count <= 3 else ""
    else:
        hdr_model = model_no
        hdr_serial = serials_str

    p_to_remove = list(c0.paragraphs)
    sub_table = c0.add_table(rows=3, cols=2)
    for p in p_to_remove:
        try:
            p._p.getparent().remove(p._p)
        except Exception:
            pass

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

    col_w = [6400, 4400]
    for row in sub_table.rows:
        for c_i, w in enumerate(col_w):
            row.cells[c_i].width = int(w * 635)
            row.cells[c_i]._tc.tcPr.tcW.w = w

    hdr_data = [
        (("Job No: ", job_no), ("Model No: ", hdr_model)),
        (("Customer: ", customer), ("Sl. No: ", hdr_serial)),
        (("Item Description: ", item_desc), ("Qty: ", qty))
    ]

    for r_idx, (col0_data, col1_data) in enumerate(hdr_data):
        # Col 0
        p0 = sub_table.rows[r_idx].cells[0].paragraphs[0]
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(col0_data[0])
        r0.bold = True
        r0.font.name = 'Times New Roman'
        r0.font.size = Pt(11)
        r0_v = p0.add_run(col0_data[1])
        r0_v.bold = True
        r0_v.font.name = 'Times New Roman'
        r0_v.font.size = Pt(11)

        # Col 1
        p1 = sub_table.rows[r_idx].cells[1].paragraphs[0]
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(col1_data[0])
        r1.bold = True
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(11)
        r1_v = p1.add_run(col1_data[1])
        r1_v.bold = True
        r1_v.font.name = 'Times New Roman'
        r1_v.font.size = Pt(11)

    # 2. Body (Table 0 Row 1 Cell 0)
    r1 = doc.tables[0].rows[1]
    c1 = r1.cells[0]
    if mode == "general":
        trPr = r1._tr.get_or_add_trPr()
        for old_th in trPr.findall(qn('w:trHeight')):
            trPr.remove(old_th)
        trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="11940" w:hRule="atLeast"/>'))
    tracking_no = get_val("tracking_no", "tracking_number", default="")
    test_readings = get_val("readings_result", "test_readings", "readings", "observations", "result", default="")

    if mode == "load_cell":
        # Tracking Number
        p_trk = c1.paragraphs[2] if len(c1.paragraphs) > 2 else None
        if p_trk:
            if tracking_no:
                p_trk.text = f"TRACKING NO: {tracking_no}"
                for r in p_trk.runs:
                    r.bold = True
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(11)
            else:
                p_trk.text = ""

        # Populate Nested Table
        if c1.tables:
            nt = c1.tables[0]
            lc_entries = []
            for idx in range(1, 9):
                s = get_val(f"lc_serial_{idx}", f"lc_sl_{idx}", default="")
                in_r = get_val(f"lc_in_res_{idx}", f"lc_rin_{idx}", default="")
                out_r = get_val(f"lc_out_res_{idx}", f"lc_rout_{idx}", default="")
                unb = get_val(f"lc_unbalance_{idx}", f"lc_zero_{idx}", default="")
                if s or in_r or out_r or unb:
                    lc_entries.append((s, in_r, out_r, unb))

            if not lc_entries and serials_list:
                for s in serials_list[:4]:
                    lc_entries.append((s, "", "", ""))

            # Fill existing rows
            for idx in range(1, len(nt.rows)):
                if idx - 1 < len(lc_entries):
                    s, in_r, out_r, unb = lc_entries[idx - 1]
                    row_vals = [str(idx), s, in_r, out_r, unb]
                else:
                    row_vals = ["", "", "", "", ""]
                cells = nt.rows[idx].cells
                for ci, val in enumerate(row_vals):
                    cells[ci].text = val
                    p = cells[ci].paragraphs[0]
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    if p.runs:
                        p.runs[0].font.name = 'Times New Roman'
                        p.runs[0].font.size = Pt(11)

        # Observations Text below table
        obs_text = test_readings.strip()
        if not obs_text:
            obs_text = "The above mentioned load cells 3 boxes (sl no, input resistance, output resistance, and initial unbalance) were checked and found OK."

        obs_lines = [l.strip() for l in obs_text.split('\n') if l.strip()]
        p_idx = 5
        for line in obs_lines:
            if p_idx < len(c1.paragraphs):
                p = c1.paragraphs[p_idx]
            else:
                p = c1.add_paragraph()
            p.text = line
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)
                if line.isupper() or line.startswith('LOAD CELL'):
                    r.bold = True
            p_idx += 1

        # Remove extra trailing empty paragraphs in c1 to preserve strictly 1 Page
        while len(c1.paragraphs) > p_idx:
            last_p = c1.paragraphs[-1]
            last_p._p.getparent().remove(last_p._p)

    elif mode == "remote_display":
        p_trk = c1.paragraphs[2] if len(c1.paragraphs) > 2 else None
        if p_trk:
            if tracking_no:
                p_trk.text = f"TRACKING NO: {tracking_no}"
                for r in p_trk.runs:
                    r.bold = True
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(11)
            else:
                p_trk.text = ""

        p_sl = c1.paragraphs[4] if len(c1.paragraphs) > 4 else None
        if p_sl:
            p_sl.text = f"Sl. No: {serials_str}" if serials_str else ""
            for r in p_sl.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)

        ind_model = get_val("indicator_model", "remote_ind_ref", "ind_model", "di_model", default="")
        p_comm = c1.paragraphs[9] if len(c1.paragraphs) > 9 else None
        if p_comm:
            if test_readings.strip():
                p_comm.text = test_readings.strip()
            else:
                rep = f"{ind_model} " if ind_model else ""
                p_comm.text = f"The above-mentioned items were checked for its communication (rs485) using {rep}indicator and was found ok"
            for r in p_comm.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)

    elif mode == "jb_single":
        p_sl = c1.paragraphs[4] if len(c1.paragraphs) > 4 else None
        if p_sl:
            p_sl.text = f"Sl. No: {serials_str}" if serials_str else ""
            for r in p_sl.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)

        p_obs = c1.paragraphs[25] if len(c1.paragraphs) > 25 else None
        if p_obs:
            if test_readings.strip():
                p_obs.text = test_readings.strip()
            else:
                p_obs.text = "The Above mentioned items were check for its continuity and was found ok"
            for r in p_obs.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)

    elif mode == "jb_multiple":
        pairs = [(2, 3), (7, 8), (11, 12), (15, 16), (19, 20)]
        jbox_entries = []
        for i in range(1, 6):
            vm = get_val(f"variant_model_{i}", f"jbox_model_{i}", default="")
            vs = get_val(f"variant_serial_{i}", f"jbox_serials_{i}", default="")
            vq = get_val(f"variant_qty_{i}", f"jbox_qty_{i}", default="")
            if not vq and vs:
                count = len([x for x in vs.split(',') if x.strip()])
                vq = f"{count} Nos" if count > 0 else ""
            if vm or vs or vq:
                jbox_entries.append((vm, vs, vq))

        for i, (pm, ps) in enumerate(pairs):
            if i < len(jbox_entries):
                vm, vs, vq = jbox_entries[i]
                set_para_tab(c1.paragraphs[pm], f"MODEL NO: {vm}", f"QTY: {vq}", 5800, 11, True)
                c1.paragraphs[ps].text = f"SL NO: {vs}"
                for r in c1.paragraphs[ps].runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(11)
            else:
                c1.paragraphs[pm].text = ""
                c1.paragraphs[ps].text = ""

        p_obs = c1.paragraphs[25] if len(c1.paragraphs) > 25 else None
        if p_obs:
            if test_readings.strip():
                p_obs.text = test_readings.strip()
            else:
                p_obs.text = "The Above mentioned items were check for its continuity and was found ok"
            for r in p_obs.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)

    elif mode == "general":
        if custom_table and custom_table.get("headers") and custom_table.get("rows"):
            headers = [str(h).strip() for h in custom_table["headers"]]
            rows = custom_table["rows"]
            num_cols = len(headers)
            num_rows = len(rows) + 1

            p0 = c1.paragraphs[0]
            p0.text = "Test Readings and Result:"
            p0.paragraph_format.space_after = Pt(6)

            tbl = c1.add_table(rows=num_rows, cols=num_cols)
            tbl.style = 'Table Grid'

            total_avail_inches = 6.8
            if num_cols == 1:
                widths = [Inches(total_avail_inches)]
            elif num_cols == 3:
                widths = [Inches(0.9), Inches(3.4), Inches(2.5)]
            elif num_cols == 4:
                widths = [Inches(0.8), Inches(2.8), Inches(1.8), Inches(1.4)]
            elif num_cols == 5:
                widths = [Inches(0.7), Inches(2.3), Inches(1.3), Inches(1.3), Inches(1.2)]
            else:
                first_w = min(1.0, total_avail_inches / num_cols)
                rem_w = (total_avail_inches - first_w) / max(1, (num_cols - 1))
                widths = [Inches(first_w)] + [Inches(rem_w)] * (num_cols - 1)

            for ci, h in enumerate(headers):
                c = tbl.rows[0].cells[ci]
                c.width = widths[min(ci, len(widths) - 1)]
                c.text = h
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E2E8F0"/>')
                c._tc.get_or_add_tcPr().append(shd)
                p = c.paragraphs[0]
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                if p.runs:
                    p.runs[0].font.name = 'Times New Roman'
                    p.runs[0].font.size = Pt(10.5)
                    p.runs[0].font.bold = True

            for r_idx, row_vals in enumerate(rows, start=1):
                for ci in range(num_cols):
                    val = str(row_vals[ci]).strip() if ci < len(row_vals) else ""
                    c = tbl.rows[r_idx].cells[ci]
                    c.width = widths[min(ci, len(widths) - 1)]
                    c.text = val
                    p = c.paragraphs[0]
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    if ci == 0 or (ci == num_cols - 1 and len(val) < 15):
                        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
                    if p.runs:
                        p.runs[0].font.name = 'Times New Roman'
                        p.runs[0].font.size = Pt(10)

            obs_text = test_readings.strip()
            if not obs_text:
                obs_text = "The above custom items and parameters were inspected, tested, and verified according to standard specifications and found OK."

            p_end = c1.add_paragraph()
            p_end.paragraph_format.space_before = Pt(8)
            p_end.text = obs_text
            for r in p_end.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)

            to_delete = [p for p in c1.paragraphs if p._p != p0._p and p._p != p_end._p]
            for p in to_delete:
                p._p.getparent().remove(p._p)

        else:
            obs_text = test_readings.strip()
            if not obs_text:
                label_lines = []
                for i in range(1, 15):
                    lbl = get_val(f"entry_label_{i}", default="")
                    val = get_val(f"entry_val_{i}", default="")
                    if lbl or val:
                        label_lines.append(f"{lbl}: {val}" if lbl else val)
                if label_lines:
                    obs_text = "\n".join(label_lines)
                else:
                    obs_text = "All items and subsystems were inspected, tested, and verified according to standard specifications and found OK."

            lines = obs_text.split('\n')
            p_idx = 1
            for line in lines:
                line_str = line.strip()
                if not line_str:
                    continue
                if p_idx < len(c1.paragraphs):
                    p = c1.paragraphs[p_idx]
                else:
                    p = c1.add_paragraph()
                p.text = line_str
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(11)
                    if line_str.isupper() or any(line_str.startswith(h) for h in ['1.', '2.', '3.', '4.', '5.', 'LOAD CELL', 'JUNCTION BOX', 'REMOTE DISPLAY', 'DIGITAL INDICATOR']):
                        r.bold = True
                p_idx += 1

            while p_idx < len(c1.paragraphs):
                c1.paragraphs[p_idx].text = ""
                p_idx += 1

    # 3. Sign-off
    tested_by = get_val("tested_by", "user_name", default="")
    approved_by = get_val("approved_by", default="") or "HOD-PDN"
    date_val = get_val("date", default="") or datetime.now().strftime("%d-%b-%Y")

    set_para_with_tab(doc.paragraphs[3], f"Tested by: {tested_by}", f"Approved by: {approved_by}", 7200, 11, bold=False)
    set_para_with_tab(doc.paragraphs[4], f"Date: {date_val}", "(HOD-PDN)", 7200, 11, bold=False)

    doc.save(output_path)
    convert_docx_to_pdf(output_path)
    return output_path

def generate_crane_scale_record(data, output_path):
    master_path = os.path.join(TEMPLATES_DIR, "PP-05_08 Crane Scale.docx")
    if not os.path.exists(master_path):
        raise FileNotFoundError(f"Template not found: {master_path}")

    doc = docx.Document(master_path)

    def get_val(*keys, default=""):
        for k in keys:
            v = data.get(k)
            if v is not None and str(v).strip():
                return str(v).strip()
        return default

    job_no = get_val("job_no", default="")
    customer = get_val("customer", "customer_name", default="")
    capacity = get_val("capacity", default="")
    stamp_no = get_val("stamp_no", default="")
    lc_model = get_val("lc_model", "load_cell_model", default="")
    lc_serial = get_val("lc_serial", "load_cell_serial", default="")
    resolution = get_val("resolution", default="")
    accuracy = get_val("accuracy", default="")

    cust_str = str(customer or "")
    if len(cust_str) > 55:
        hdr_tab, hdr_pt = 3600, 9.5
    elif len(cust_str) > 40:
        hdr_tab, hdr_pt = 3900, 10.0
    else:
        hdr_tab, hdr_pt = 4200, 11.0

    if len(doc.paragraphs) > 2:
        set_para_with_tab(doc.paragraphs[2], f"        JOB No.    : {job_no}", f"CUSTOMER: {customer}", hdr_tab, hdr_pt, bold=True)
    if len(doc.paragraphs) > 3:
        set_para_with_tab(doc.paragraphs[3], f"        CAPACITY  : {capacity}", f"STAMP NO: {stamp_no}", hdr_tab, hdr_pt, bold=True)
    if len(doc.paragraphs) > 4:
        set_para_with_tab(doc.paragraphs[4], f"        Load cell. Model No. : {lc_model}", f"Serial No. : {lc_serial}", hdr_tab, hdr_pt, bold=True)
    if len(doc.paragraphs) > 8:
        res_text = f"\t   Resolution :\tL: {resolution}" if resolution else "\t   Resolution :\tL:"
        if accuracy:
            res_text += f"     Accuracy: {accuracy}"
        doc.paragraphs[8].text = res_text
        for r in doc.paragraphs[8].runs:
            r.font.name = 'Times New Roman'
            r.font.size = Pt(11)

    # Hardware Table (Table 0)
    crane_model = get_val("crane_model", default="")
    crane_serial = get_val("crane_serial", default="")
    charger_model = get_val("charger_model", "charger_details", default="")
    charger_serial = get_val("charger_serial", default="")
    remote_model = get_val("remote_model", "remote_details", default="")
    remote_serial = get_val("remote_serial", default="")
    hook_serial = get_val("hook_serial", default="")
    shackle_serial = get_val("shackle_serial", default="")

    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        if len(t0.rows) > 1:
            r1 = t0.rows[1]
            r1.cells[0].text = f"Model: {crane_model}\nSl. No: {crane_serial}".strip()
            r1.cells[1].text = f"{charger_model}\nSl. No: {charger_serial}" if charger_serial and charger_serial not in charger_model else charger_model
            r1.cells[2].text = f"{remote_model}\nSl. No: {remote_serial}" if remote_serial and remote_serial not in remote_model else remote_model
            r1.cells[3].text = f"Sl. No: {hook_serial}" if (hook_serial and not hook_serial.startswith("Sl.")) else (hook_serial or "Sl. No:")
            r1.cells[4].text = f"Sl. No: {shackle_serial}" if (shackle_serial and not shackle_serial.startswith("Sl.")) else (shackle_serial or "Sl. No:")
            for cell in r1.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(11)

    # Routine Test Table (Table 1) - Guarded against AC power / 24V form cross-contamination
    raw_act1 = get_val("act_1", "ac1", default="")
    if any(x in str(raw_act1).upper() for x in ["230", "AC", "DERIVED", "LINE"]):
        act_1 = ""
    else:
        act_1 = raw_act1

    raw_spec2 = get_val("spec_2", "sp2", default="")
    raw_act2 = get_val("act_2", "ac2", default="")
    spec_2 = "" if any(x in str(raw_spec2).upper() for x in ["24", "DC", "VOLT"]) else raw_spec2
    act_2 = "" if any(x in str(raw_act2).upper() for x in ["24", "DC", "VOLT"]) else raw_act2

    act_3 = get_val("act_3", "ac3", default="")
    spec_3 = get_val("spec_3", "sp3", default="")

    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        if len(t1.rows) > 3:
            if act_1: t1.rows[1].cells[3].text = act_1
            if spec_2: t1.rows[2].cells[2].text = spec_2
            if act_2: t1.rows[2].cells[3].text = act_2
            if spec_3: t1.rows[3].cells[2].text = spec_3
            if act_3: t1.rows[3].cells[3].text = act_3
            for r_idx in range(1, 4):
                for cell in t1.rows[r_idx].cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.name = 'Times New Roman'
                            r.font.size = Pt(11)

    # Battery & Remote checklist paragraphs
    batt_conn = get_val("batt_connected", "battery_connected", default="")
    low_batt = get_val("low_batt_voltage", "low_batt", default="")
    remote_op = get_val("remote_op", "remote_operation", default="")
    inst_used = get_val("instrument_used", default="")
    remarks = get_val("remarks", default="")

    if len(doc.paragraphs) > 13 and batt_conn:
        doc.paragraphs[13].text = f"        When the battery connected: {batt_conn}"
        for r in doc.paragraphs[13].runs: r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    if len(doc.paragraphs) > 15 and low_batt:
        doc.paragraphs[15].text = f"Low battery voltage set (CS): {low_batt}"
        for r in doc.paragraphs[15].runs: r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    if len(doc.paragraphs) > 17 and remote_op:
        doc.paragraphs[17].text = f"Check for the remote operation: {remote_op}"
        for r in doc.paragraphs[17].runs: r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    if len(doc.paragraphs) > 22 and remarks:
        doc.paragraphs[22].text = f"Remarks (if any): {remarks}"
        for r in doc.paragraphs[22].runs: r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    if len(doc.paragraphs) > 29:
        doc.paragraphs[29].text = f"                 Instrument Used: {inst_used}"
        for r in doc.paragraphs[29].runs: r.font.name = 'Times New Roman'; r.font.size = Pt(11)

    tested_by = get_val("tested_by", "user_name", default="")
    approved_by = get_val("approved_by", default="")
    date_val = get_val("date", default="") or datetime.now().strftime("%d-%b-%Y")

    if len(doc.paragraphs) > 31:
        set_para_with_tab(doc.paragraphs[31], f"              Checked by: {tested_by}", f"Approved by: {approved_by}", 6500, 11, bold=False)
    if len(doc.paragraphs) > 32:
        set_para_with_tab(doc.paragraphs[32], f"\t       Date: {date_val}", "(HOD - PDN)", 6500, 11, bold=False)

    # Remove 1 excess empty spacer paragraph between Remarks and Instrument Used to ensure strictly 1 Page
    for i, p in enumerate(doc.paragraphs):
        if "Remarks" in p.text:
            for j in range(i + 1, len(doc.paragraphs)):
                if "Instrument Used" in doc.paragraphs[j].text:
                    break
                if not doc.paragraphs[j].text.strip():
                    try:
                        doc.paragraphs[j]._p.getparent().remove(doc.paragraphs[j]._p)
                    except Exception:
                        pass
                    break
            break

    doc.save(output_path)
    convert_docx_to_pdf(output_path)
    return output_path


def _gval(data, field_map, *keys, default=""):
    """Return the first non-empty string value from data/field_map for the given candidate keys."""
    for k in keys:
        for src in (field_map, data):
            if src and k in src:
                v = src.get(k)
                if v is not None and str(v).strip() != "":
                    return str(v).strip()
    return default


def _style_para(p, pt=10.5, name="Times New Roman", bold=None):
    for r in p.runs:
        r.font.name = name
        r.font.size = Pt(pt)
        if bold is not None:
            r.bold = bold


def _set_para_text(p, text, pt=10.5, bold=None):
    p.text = "" if text is None else str(text)
    _style_para(p, pt=pt, bold=bold)


def _set_cell_text(cell, text, pt=10.0, bold=None):
    cell.text = "" if text is None else str(text)
    for p in cell.paragraphs:
        _style_para(p, pt=pt, bold=bold)


def _unique_row_cells(row):
    """Return unique CT_Tc cell objects in left-to-right visual order (handling merged spans safely)."""
    seen = set()
    uniq = []
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id not in seen:
            seen.add(tc_id)
            uniq.append(cell)
    return uniq


def _fill_cell_ordered_labels(cell, rules, pt=10.0):
    """Fill sequential labeled lines inside a table cell using an ordered list of (matcher, formatter).

    Each tuple is (predicate_or_substring, value_str, replacement_line_fn_or_None).
    Matching consumes the rule once so repeated labels (e.g. multiple Model No. / Serial No.) stay distinct.
    """
    queue = list(rules)
    for p in cell.paragraphs:
        txt = p.text
        lines = txt.split("\n")
        changed = False
        new_lines = []
        for ln in lines:
            stripped = ln.strip()
            matched_idx = None
            for idx, rule in enumerate(queue):
                matcher = rule[0]
                if callable(matcher):
                    ok = matcher(stripped)
                else:
                    ok = stripped.lower().startswith(matcher.lower())
                if ok:
                    matched_idx = idx
                    break
            if matched_idx is not None:
                _, val, fmt_fn = queue.pop(matched_idx)
                val_s = "" if val is None else str(val).strip()
                if fmt_fn is not None:
                    new_lines.append(fmt_fn(stripped, val_s))
                elif ":" in stripped:
                    lbl = stripped.split(":", 1)[0].strip()
                    new_lines.append(f"{lbl}: {val_s}" if val_s else f"{lbl}:")
                else:
                    new_lines.append(f"{stripped}: {val_s}" if val_s else stripped)
                changed = True
            else:
                new_lines.append(ln)
        if changed:
            _set_para_text(p, "\n".join(new_lines), pt=pt)


def _fill_paras_by_prefix(paragraphs, rules, pt=10.5):
    """Consume (prefix_tuple, formatter) rules in document order against paragraphs."""
    queue = list(rules)
    for p in paragraphs:
        stripped = p.text.strip()
        if not stripped:
            continue
        for idx, (prefixes, fn) in enumerate(queue):
            if isinstance(prefixes, str):
                prefixes = (prefixes,)
            if any(stripped.lower().startswith(pref.lower()) for pref in prefixes):
                fn(p, stripped)
                queue.pop(idx)
                break


def _fill_job_traveller_card(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "customer_name", "cust")
    sys_name = _gval(data, field_map, "system_name", "system", "model_no")
    qty = _gval(data, field_map, "qty")
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.startswith("JOB No"):
            set_para_with_tab(p, f"\tJOB No: {job}", f"CUSTOMER: {cust}", 4800, 11.0, bold=False)
        elif txt.startswith("SYSTEM"):
            set_para_with_tab(p, f"\tSYSTEM: {sys_name}", f"QTY: {qty}", 4800, 11.0, bold=False)

    if not doc.tables:
        return
    t0 = doc.tables[0]
    if len(t0.rows) > 0:
        d_sig = _gval(data, field_map, "design_sig")
        d_dt = _gval(data, field_map, "design_date")
        d_nm = _gval(data, field_map, "design_name")
        _fill_cell_ordered_labels(t0.rows[0].cells[0], [
            ("Signature", f"{d_sig}\t\tDate: {d_dt}".strip() if d_dt else d_sig,
             lambda _, v: f"Signature: {d_sig:<45} Date: {d_dt}"),
            ("Name", d_nm, None),
        ], pt=10.0)

    if len(t0.rows) > 3:
        _fill_cell_ordered_labels(t0.rows[1].cells[0], [
            ("Doc. No", _gval(data, field_map, "asm_doc_no"), None),
            ("NAME", _gval(data, field_map, "asm_chassis_name"), None),
            ("Starting date", _gval(data, field_map, "asm_chassis_start"), None),
            ("Completion date", _gval(data, field_map, "asm_chassis_end"), None),
            ("Signature", _gval(data, field_map, "asm_chassis_sig"), None),
            ("NAME", _gval(data, field_map, "asm_final_name"), None),
            ("Starting date", _gval(data, field_map, "asm_final_start"), None),
            ("Completion date", _gval(data, field_map, "asm_final_end"), None),
            ("Feedback", _gval(data, field_map, "asm_final_feedback"), None),
            ("Signature", _gval(data, field_map, "asm_final_sig"), None),
            ("Inspected by", _gval(data, field_map, "asm_inspected_by"), None),
        ], pt=9.5)
        rew_asm = _gval(data, field_map, "rework_asm_reason")
        if rew_asm:
            _set_cell_text(t0.rows[2].cells[1], f"Reasons and Nature of Rework:\n{rew_asm}", pt=9.5)
        _fill_cell_ordered_labels(t0.rows[3].cells[2], [
            (lambda s: s == "By:", _gval(data, field_map, "rework_asm_by"), lambda _, v: f"Carried out By: {v}"),
            (lambda s: s == "Taken:", _gval(data, field_map, "rework_asm_time"), lambda _, v: f"Time Taken: {v}"),
            ("Inspected", _gval(data, field_map, "rework_asm_inspected"), None),
        ], pt=9.5)

    if len(t0.rows) > 4:
        _fill_cell_ordered_labels(t0.rows[4].cells[0], [
            ("NAME", _gval(data, field_map, "wiring_name"), None),
            ("Doc. No", _gval(data, field_map, "wiring_doc_no"), None),
            ("Starting date", _gval(data, field_map, "wiring_start"), None),
            ("Completion date", _gval(data, field_map, "wiring_end"), None),
            ("Feedback", _gval(data, field_map, "wiring_feedback"), None),
            ("Signature", _gval(data, field_map, "wiring_sig"), None),
            ("Inspected by", _gval(data, field_map, "wiring_inspected_by"), None),
        ], pt=9.5)
        rew_w = _gval(data, field_map, "rework_wiring_reason")
        if rew_w:
            _set_cell_text(t0.rows[4].cells[1], rew_w, pt=9.5)
        _fill_cell_ordered_labels(t0.rows[4].cells[2], [
            (lambda s: s == "By:", _gval(data, field_map, "rework_wiring_by"), lambda _, v: f"Carried out By: {v}"),
            (lambda s: s == "Taken:", _gval(data, field_map, "rework_wiring_time"), lambda _, v: f"Time Taken: {v}"),
            ("Inspected", _gval(data, field_map, "rework_wiring_inspected"), None),
        ], pt=9.5)

    if len(t0.rows) > 5:
        _fill_cell_ordered_labels(t0.rows[5].cells[0], [
            ("NAME", _gval(data, field_map, "testing_name"), None),
            ("Doc. No", _gval(data, field_map, "testing_doc_no"), None),
            ("Starting date", _gval(data, field_map, "testing_start"), None),
            ("Completion date", _gval(data, field_map, "testing_end"), None),
            ("Given to Q.C.", _gval(data, field_map, "testing_qc_date"), None),
            ("Feedback", _gval(data, field_map, "testing_feedback"), None),
            ("Signature", _gval(data, field_map, "testing_sig"), None),
        ], pt=9.5)
        rew_t = _gval(data, field_map, "rework_testing_reason")
        if rew_t:
            _set_cell_text(t0.rows[5].cells[1], rew_t, pt=9.5)
        _fill_cell_ordered_labels(t0.rows[5].cells[2], [
            (lambda s: s == "By:", _gval(data, field_map, "rework_testing_by"), lambda _, v: f"Carried out By: {v}"),
            (lambda s: s == "Taken:", _gval(data, field_map, "rework_testing_time"), lambda _, v: f"Time Taken: {v}"),
            ("Inspected", _gval(data, field_map, "rework_testing_inspected"), None),
        ], pt=9.5)


def _fill_batching_system(doc, data, field_map):
    res_l = _gval(data, field_map, "res_l", "rl")
    inst = _gval(data, field_map, "instrument_used", "inst")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("Resolution:", lambda p, _: _set_para_text(p, f"Resolution: L: {res_l} kg/T", pt=10.5)),
        ("Instrument Used:", lambda p, _: _set_para_text(p, f"      Instrument Used: {inst}", pt=10.5)),
        ("Tested by:", lambda p, _: set_para_with_tab(p, f"\tTested by: {tb}", f"Approved by: {ab}", 6200, 10.5, bold=False)),
        ("Date:", lambda p, _: set_para_with_tab(p, f"      Date: {dt}", "(HOD – PDN)", 6200, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        r1 = doc.tables[0].rows[1]
        _fill_cell_ordered_labels(r1.cells[0], [
            ("Model No.", _gval(data, field_map, "bch_model", "model_no"), None),
            ("Serial No.", _gval(data, field_map, "bch_serial", "serial_no"), None),
        ], pt=9.5)
        _fill_cell_ordered_labels(r1.cells[1], [
            ("Model No.", _gval(data, field_map, "ctrl_model"), None),
            ("Serial No.", _gval(data, field_map, "ctrl_serial"), None),
            ("SW version", _gval(data, field_map, "sw_version"), None),
        ], pt=9.5)
        _fill_cell_ordered_labels(r1.cells[2], [
            ("Model No.", _gval(data, field_map, "lc_model", "sensor_model"), None),
            ("1)", _gval(data, field_map, "s1"), lambda _, v: f"1) {v}"),
            ("2)", _gval(data, field_map, "s2"), lambda _, v: f"2) {v}"),
            ("3)", _gval(data, field_map, "s3"), lambda _, v: f"3) {v}"),
            ("4)", _gval(data, field_map, "s4"), lambda _, v: f"4) {v}"),
        ], pt=9.5)
        _fill_cell_ordered_labels(r1.cells[3], [
            ("Model No.", _gval(data, field_map, "jbox_model1", "jb_1_model"), None),
            ("Serial No.", _gval(data, field_map, "jbox_serial1", "jb_1_serial"), None),
            ("Model No.", _gval(data, field_map, "jbox_model2", "jb_2_model"), None),
            ("Serial No.", _gval(data, field_map, "jbox_serial2", "jb_2_serial"), None),
        ], pt=9.5)

    sc = _gval(data, field_map, "short_check", "test_voltage_chk")
    if sc and len(doc.tables) > 1:
        _set_cell_text(doc.tables[1].rows[0].cells[0],
                       f"NOTE: Check for short between Line & Neutral, Neutral & Earth, Line & Earth: {sc}", pt=10.0)

    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for i in range(1, min(6, len(t2.rows))):
            sp = _gval(data, field_map, f"spec_{i}", f"sp{i}")
            ac = _gval(data, field_map, f"act_{i}", f"ac{i}")
            if sp: _set_cell_text(t2.rows[i].cells[2], sp, pt=10.0)
            if ac: _set_cell_text(t2.rows[i].cells[3], ac, pt=10.0)

    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        for i in range(1, min(7, len(t3.rows))):
            for c in range(3):
                v = _gval(data, field_map, f"g{i}_{c}")
                if v: _set_cell_text(t3.rows[i].cells[c], v, pt=10.0)

    if len(doc.tables) > 4:
        t4 = doc.tables[4]
        for i in range(1, min(11, len(t4.rows))):
            v = _gval(data, field_map, f"func_{i}")
            if v: _set_cell_text(t4.rows[i].cells[2], v, pt=10.0)
        rem = _gval(data, field_map, "remarks")
        if len(t4.rows) > 11 and rem:
            _set_cell_text(t4.rows[11].cells[0], f"REMARKS (IF ANY): {rem}", pt=10.0)


def _fill_remote_indicator(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    cap = _gval(data, field_map, "capacity", "cap")
    rate_v = _gval(data, field_map, "rate_val")
    load_v = _gval(data, field_map, "load_val")
    speed_v = _gval(data, field_map, "speed_val")
    tot_v = _gval(data, field_map, "tot_val")
    inst = _gval(data, field_map, "instrument_used")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("JOB No", lambda p, _: set_para_with_tab(p, f"JOB No. : {job}", f"CUSTOMER: {cust}", 4500, 10.5, bold=False)),
        ("CAPACITY", lambda p, _: _set_para_text(p, f"CAPACITY : {cap}", pt=10.5)),
        ("RATE:", lambda p, _: set_para_with_tab(p, f"RATE: {rate_v}", f"LOAD: {load_v}", 4500, 10.5, bold=False)),
        ("SPEED:", lambda p, _: set_para_with_tab(p, f"SPEED: {speed_v}", f"TOTALIZER: {tot_v}", 4500, 10.5, bold=False)),
        ("Instrument Used:", lambda p, _: _set_para_text(p, f"    Instrument Used: {inst}", pt=10.5)),
        ("Tested by:", lambda p, _: set_para_with_tab(p, f"      Tested by: {tb}", f"Approved by: {ab}", 6200, 10.5, bold=False)),
        ("Date:", lambda p, _: set_para_with_tab(p, f"     Date: {dt}", "(HOD – PDN)", 6200, 10.5, bold=False)),
    ])

    sc = _gval(data, field_map, "short_check", "test_voltage_chk")
    if sc and len(doc.tables) > 0:
        _set_cell_text(doc.tables[0].rows[0].cells[0],
                       f"Check for short between Line & Neutral, Neutral & Earth, Line & Earth : {sc}", pt=10.0)

    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for i in range(1, min(10, len(t1.rows))):
            sp = _gval(data, field_map, f"spec_{i}", f"sp{i}")
            ac = _gval(data, field_map, f"act_{i}", f"ac{i}")
            if sp: _set_cell_text(t1.rows[i].cells[2], sp, pt=10.0)
            if ac: _set_cell_text(t1.rows[i].cells[3], ac, pt=10.0)


def _fill_weigh_feeder(doc, data, field_map):
    rl = _gval(data, field_map, "res_load", "res_l")
    rs = _gval(data, field_map, "res_speed", "res_s")
    rr = _gval(data, field_map, "res_rate", "res_r")
    rt = _gval(data, field_map, "res_tot", "res_t")
    cw = _gval(data, field_map, "cal_chain_wt")
    vm = _gval(data, field_map, "vol_mode")
    gm = _gval(data, field_map, "grav_mode")
    rem = _gval(data, field_map, "remarks")
    inst = _gval(data, field_map, "instrument_used")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("Resolution:", lambda p, _: _set_para_text(p, f"Resolution: Load: {rl} kg/m   Speed: {rs} m/s   Rate: {rr} tph   Totalizer: {rt} tonnes", pt=10.0)),
        ("III).", lambda p, _: _set_para_text(p, f"III). Load Calibration:   Weight of Calibration Chain: {cw} Kg/m", pt=10.0)),
        ("IV).", lambda p, _: _set_para_text(p, f"IV). VOLUMETRIC MODE: {vm}", pt=10.0)),
        ("V).", lambda p, _: _set_para_text(p, f"V). GRAVIMETRIC MODE: Operating Range: {gm}", pt=10.0)),
        ("Remarks:", lambda p, _: _set_para_text(p, f"Remarks: {rem}", pt=10.0)),
        ("Instrument Used:", lambda p, _: _set_para_text(p, f" Instrument Used: {inst}", pt=10.0)),
        ("Tested by:", lambda p, _: set_para_with_tab(p, f"\tTested by: {tb}", f"Approved by: {ab}", 6200, 10.5, bold=False)),
        ("Date:", lambda p, _: set_para_with_tab(p, f"\tDate: {dt}", "(HOD – PDN)", 6200, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        r1 = doc.tables[0].rows[1]
        _fill_cell_ordered_labels(r1.cells[0], [
            ("Model No.", _gval(data, field_map, "wf_model", "model_no"), None),
            ("Serial No.", _gval(data, field_map, "wf_serial", "serial_no"), None),
            ("Model No.", _gval(data, field_map, "lp_model"), None),
            ("Serial No.", _gval(data, field_map, "lp_serial"), None),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[1], [
            ("Model No.", _gval(data, field_map, "ctrl_model"), None),
            ("Serial No.", _gval(data, field_map, "ctrl_serial"), None),
            ("SW version", _gval(data, field_map, "ctrl_sw", "sw_version"), None),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[2], [
            ("Model No.", _gval(data, field_map, "lc_model"), None),
            ("1)", _gval(data, field_map, "lc_serial_1", "s1"), lambda _, v: f"1) {v}"),
            ("2)", _gval(data, field_map, "lc_serial_2", "s2"), lambda _, v: f"2) {v}"),
            ("3)", _gval(data, field_map, "lc_serial_3", "s3"), lambda _, v: f"3) {v}"),
            ("4)", _gval(data, field_map, "lc_serial_4", "s4"), lambda _, v: f"4) {v}"),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[3], [
            ("Model No.", _gval(data, field_map, "tacho_model", "ss_model"), None),
            ("Serial No.", _gval(data, field_map, "tacho_serial", "ss_serial"), None),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[4], [
            ("Model No.", _gval(data, field_map, "jbox_1_model", "jbox_model1"), None),
            ("Serial No.", _gval(data, field_map, "jbox_1_serial", "jbox_serial1"), None),
            ("Model No.", _gval(data, field_map, "jbox_2_model", "jbox_model2"), None),
            ("Serial No.", _gval(data, field_map, "jbox_2_serial", "jbox_serial2"), None),
            ("Model No.", _gval(data, field_map, "jbox_3_model", "jbox_model3"), None),
            ("Serial No.", _gval(data, field_map, "jbox_3_serial", "jbox_serial3"), None),
        ], pt=9.0)

    if len(doc.tables) > 1 and len(doc.tables[1].rows) > 1:
        r1 = doc.tables[1].rows[1]
        _fill_cell_ordered_labels(r1.cells[0], [
            ("Make", _gval(data, field_map, "drive_make"), None),
            ("Model", _gval(data, field_map, "drive_model"), None),
            ("Serial", _gval(data, field_map, "drive_serial"), None),
            ("Capacity", _gval(data, field_map, "drive_kw"), lambda _, v: f"Capacity: {v} kW" if v else "Capacity: kW"),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[1], [
            ("Make", _gval(data, field_map, "motor_make"), None),
            ("Model No.", _gval(data, field_map, "motor_model"), None),
            ("Serial", _gval(data, field_map, "motor_serial"), None),
            ("Rating", _gval(data, field_map, "motor_kw"), lambda _, v: f"Rating: {v} kW" if v else "Rating: kW"),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[2], [
            ("Make", _gval(data, field_map, "gb_make"), None),
            ("Ratio", _gval(data, field_map, "gb_ratio"), None),
            ("Model No.", _gval(data, field_map, "gb_model"), None),
            ("Serial", _gval(data, field_map, "gb_serial"), None),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[3], [
            ("Make", _gval(data, field_map, "ral_make"), None),
            ("Model No.", _gval(data, field_map, "ral_model"), None),
            ("Serial", _gval(data, field_map, "ral_serial"), None),
        ], pt=9.0)
        sc = _gval(data, field_map, "short_check", "test_voltage_chk")
        if sc and len(doc.tables[1].rows) > 2:
            _set_cell_text(doc.tables[1].rows[2].cells[0],
                           f"NOTE: Check for short between Line & Neutral, Neutral & Earth, Line & Earth: {sc}", pt=9.5)

    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for i in range(1, min(6, len(t2.rows))):
            sp = _gval(data, field_map, f"spec_{i}")
            ac = _gval(data, field_map, f"act_{i}")
            if sp: _set_cell_text(t2.rows[i].cells[2], sp, pt=9.5)
            if ac:
                if i == 4 and "kg/m" not in ac.lower(): ac = f"{ac} kg/m"
                elif i == 5 and "m/s" not in ac.lower(): ac = f"{ac} m/s"
                _set_cell_text(t2.rows[i].cells[3], ac, pt=9.5)

    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        blw = _gval(data, field_map, "belt_len_width")
        asm = _gval(data, field_map, "act_speed_measured")
        if blw and len(t3.rows) > 0:
            _set_cell_text(t3.rows[0].cells[1], f"{blw} mm." if "mm" not in blw.lower() else blw, pt=9.5)
        if asm and len(t3.rows) > 1:
            _set_cell_text(t3.rows[1].cells[1], f"{asm} m/s" if "m/s" not in asm.lower() else asm, pt=9.5)

    pcts = ("10", "25", "50", "75", "100")
    if len(doc.tables) > 4:
        t4 = doc.tables[4]
        for idx, pct in enumerate(pcts, start=1):
            if idx < len(t4.rows):
                cs = _gval(data, field_map, f"calc_speed_{pct}")
                ms = _gval(data, field_map, f"meas_speed_{pct}")
                if cs: _set_cell_text(t4.rows[idx].cells[1], cs, pt=9.5)
                if ms: _set_cell_text(t4.rows[idx].cells[2], ms, pt=9.5)

    if len(doc.tables) > 5:
        t5 = doc.tables[5]
        cols_keys = ("set_tph", "act_avg", "load_kgm", "tot_act", "tot_meas", "pct_err", "ma_out")
        for idx, pct in enumerate(pcts, start=2):
            if idx < len(t5.rows):
                for c_idx, pref in enumerate(cols_keys, start=1):
                    v = _gval(data, field_map, f"{pref}_{pct}")
                    if v: _set_cell_text(t5.rows[idx].cells[c_idx], v, pt=9.5)

    if len(doc.tables) > 6:
        t6 = doc.tables[6]
        for r_i, fk in enumerate(("func_mode", "func_di", "func_pf", "func_aio", "func_comm")):
            v = _gval(data, field_map, fk)
            if v and r_i < len(t6.rows):
                _set_cell_text(t6.rows[r_i].cells[1], v, pt=9.5)


def _fill_trip_safe(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    ts_m = _gval(data, field_map, "ts_model", "model_no")
    ts_s = _gval(data, field_map, "ts_serial", "serial_no")
    lc_m = _gval(data, field_map, "lc_model")
    lc_s = _gval(data, field_map, "lc_serial")
    trip_err = _gval(data, field_map, "trip_error")
    pf_c = _gval(data, field_map, "pf_contact")
    led_a = _gval(data, field_map, "led_annunciation")
    inst = _gval(data, field_map, "instrument_used")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("JOB ORDER No.", lambda p, _: _set_para_text(p, f"\tJOB ORDER No.\t: {job}", pt=10.5)),
        ("CUSTOMER:", lambda p, _: _set_para_text(p, f"\tCUSTOMER: {cust}", pt=10.5)),
        ("MODEL NO.:", lambda p, _: set_para_with_tab(p, f"\tMODEL NO.: {ts_m}", f"MODEL NO.: {lc_m}", 5200, 10.5, bold=False)),
        ("SERIAL NO.:", lambda p, _: set_para_with_tab(p, f"\tSERIAL NO.: {ts_s}", f"SERIAL NO.: {lc_s}", 5200, 10.5, bold=False)),
        ("Percentage of error", lambda p, _: _set_para_text(p, f"               Percentage of error for the Trip Level: {trip_err}", pt=10.5)),
        ("PF Contact", lambda p, _: _set_para_text(p, f"\t\t  PF Contact for Healthy / Over Load: {pf_c}", pt=10.5)),
        ("Lamp indication", lambda p, _: _set_para_text(p, f"\t\t  Lamp indication / LED annunciation as applicable: {led_a}", pt=10.5)),
        ("Instrument Used:", lambda p, _: _set_para_text(p, f"              Instrument Used: {inst}", pt=10.5)),
        ("Tested by:", lambda p, _: set_para_with_tab(p, f"\tTested by: {tb}", f"Approved by: {ab}", 6200, 10.5, bold=False)),
        ("Date:", lambda p, _: set_para_with_tab(p, f"              Date: {dt}", "(HOD – PDN)", 6200, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        for i in range(1, 7):
            r_idx = i + 1
            if r_idx < len(t0.rows):
                vals = (
                    _gval(data, field_map, f"load_kg_{i}"),
                    _gval(data, field_map, f"input_mv_{i}"),
                    _gval(data, field_map, f"output_mv_{i}"),
                    _gval(data, field_map, f"led_level_{i}"),
                )
                for c_idx, v in enumerate(vals, start=1):
                    if v: _set_cell_text(t0.rows[r_idx].cells[c_idx], v, pt=10.0)


def _fill_vibration_switch(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    mod = _gval(data, field_map, "model_no", "vs_model")
    ser = _gval(data, field_map, "serial_no", "vs_serial")
    acc_m = _gval(data, field_map, "acc_model")
    acc_s = _gval(data, field_map, "acc_serial")
    pct_err = _gval(data, field_map, "pct_error")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("PERCENTAGE OF ERROR:", lambda p, _: _set_para_text(p, f"PERCENTAGE OF ERROR: {pct_err}", pt=10.5)),
        ("TESTED BY:", lambda p, _: set_para_with_tab(p, f"TESTED BY: {tb}", f"APPROVED BY: {ab}", 6200, 10.5, bold=False)),
        ("DATE:", lambda p, _: set_para_with_tab(p, f"DATE: {dt}", "HOD (PDN)", 6200, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        t0 = doc.tables[0]
        _fill_cell_ordered_labels(t0.rows[0].cells[0], [
            ("JOB ORDER", job, None),
            ("CUSTOMER", cust, None),
        ], pt=10.0)
        _fill_cell_ordered_labels(t0.rows[1].cells[0], [
            ("MODEL NO.", mod, None),
            ("SERIAL NO.", ser, None),
        ], pt=10.0)
        _fill_cell_ordered_labels(t0.rows[1].cells[1], [
            ("MODEL NO.", acc_m, None),
            ("SERIAL NO.", acc_s, None),
        ], pt=10.0)

    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for i in range(1, min(5, len(t1.rows))):
            in_v = _gval(data, field_map, f"cal_in_{i}", f"spec_{i}")
            out_v = _gval(data, field_map, f"cal_out_{i}", f"act_{i}")
            if in_v: _set_cell_text(t1.rows[i].cells[1], in_v, pt=10.0)
            if out_v: _set_cell_text(t1.rows[i].cells[2], out_v, pt=10.0)

    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for i in range(1, min(5, len(t2.rows))):
            for c_off in range(3):
                v = _gval(data, field_map, f"freq_trip_{i}_{c_off + 1}")
                if v: _set_cell_text(t2.rows[i].cells[5 + c_off], v, pt=10.0)


def _fill_acc_mv(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    mod = _gval(data, field_map, "model_no")
    ser = _gval(data, field_map, "serial_no")

    lin_sens_vals = []
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        for i in range(1, min(11, len(t0.rows))):
            out_v = _gval(data, field_map, f"cal_in_{i}", f"out_{i}")
            sens_v = _gval(data, field_map, f"cal_out_{i}", f"sens_{i}")
            if not sens_v and out_v:
                f_out = _parse_first_float(out_v)
                if f_out is not None:
                    sens_v = f"{f_out / i:.2f}"
            if out_v: _set_cell_text(t0.rows[i].cells[1], out_v, pt=10.0)
            if sens_v:
                _set_cell_text(t0.rows[i].cells[2], sens_v, pt=10.0)
                f_s = _parse_first_float(sens_v)
                if f_s is not None: lin_sens_vals.append(f_s)

    freq_sens_vals = []
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for i in range(1, min(7, len(t1.rows))):
            out_v = _gval(data, field_map, f"spec_{i}", f"freq_out_{i}")
            sens_v = _gval(data, field_map, f"act_{i}", f"freq_sens_{i}")
            if out_v: _set_cell_text(t1.rows[i].cells[2], out_v, pt=10.0)
            if sens_v:
                _set_cell_text(t1.rows[i].cells[3], sens_v, pt=10.0)
                f_s = _parse_first_float(sens_v)
                if f_s is not None: freq_sens_vals.append(f_s)

    post_sens_vals = []
    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for i in range(1, min(11, len(t2.rows))):
            out_v = _gval(data, field_map, f"post_out_{i}")
            sens_v = _gval(data, field_map, f"post_sens_{i}")
            if out_v: _set_cell_text(t2.rows[i].cells[1], out_v, pt=10.0)
            if sens_v:
                _set_cell_text(t2.rows[i].cells[2], sens_v, pt=10.0)
                f_s = _parse_first_float(sens_v)
                if f_s is not None: post_sens_vals.append(f_s)

    mean_lin = _gval(data, field_map, "mean_lin") or (f"{sum(lin_sens_vals)/len(lin_sens_vals):.2f}" if lin_sens_vals else "")
    mean_freq = _gval(data, field_map, "mean_freq") or (f"{sum(freq_sens_vals)/len(freq_sens_vals):.2f}" if freq_sens_vals else "")
    mean_post = _gval(data, field_map, "mean_post") or (f"{sum(post_sens_vals)/len(post_sens_vals):.2f}" if post_sens_vals else "")

    rules = [
        ("TEST REPORT FOR ACCELEROMETER:", lambda p, _: _set_para_text(p, f"TEST REPORT FOR ACCELEROMETER: MODEL NO {mod}", pt=11.0, bold=True) if mod else None),
        ("Job No:", lambda p, _: _set_para_text(p, f"Job No: {job}", pt=10.5)),
        ("Customer Name:", lambda p, _: _set_para_text(p, f"Customer Name: {cust}", pt=10.5)),
        ("Serial No:", lambda p, _: _set_para_text(p, f"Serial No: {ser}", pt=10.5)),
        ("Mean:", lambda p, _: _set_para_text(p, f"Mean: {mean_lin} mV/g", pt=10.5)),
        ("Mean:", lambda p, _: _set_para_text(p, f"Mean: {mean_freq} mV/g", pt=10.5)),
        ("Mean:", lambda p, _: _set_para_text(p, f"Mean: {mean_post} mV/g", pt=10.5)),
    ]
    _fill_paras_by_prefix(doc.paragraphs, rules)

    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        vals = (
            _gval(data, field_map, "charge_sens"),
            _gval(data, field_map, "capacitance"),
            _gval(data, field_map, "insulation_res"),
        )
        for r_idx, v in enumerate(vals):
            if r_idx < len(t3.rows) and v:
                _set_cell_text(t3.rows[r_idx].cells[1], f": {v}", pt=10.0)

    if len(doc.tables) > 4:
        t4 = doc.tables[4]
        tb = _gval(data, field_map, "tested_by", "user_name")
        ab = _gval(data, field_map, "approved_by")
        dt = _gval(data, field_map, "date")
        if len(t4.rows) > 0:
            _set_cell_text(t4.rows[0].cells[1], f": {tb}", pt=10.0)
            _set_cell_text(t4.rows[0].cells[2], f"APPROVED BY: {ab}", pt=10.0)
        if len(t4.rows) > 1:
            _set_cell_text(t4.rows[1].cells[1], f": {dt}", pt=10.0)


def _fill_acc_charge(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    mod = _gval(data, field_map, "model_no")
    ser = _gval(data, field_map, "serial_no")

    lin_sens_vals = []
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        for i in range(1, min(11, len(t0.rows))):
            amp_v = _gval(data, field_map, f"cal_in_{i}", f"amp_out_{i}")
            chg_v = _gval(data, field_map, f"charge_out_{i}")
            sens_v = _gval(data, field_map, f"cal_out_{i}", f"sens_{i}")
            f_amp = _parse_first_float(amp_v)
            if not chg_v and f_amp is not None:
                chg_v = f"{f_amp / 1.022:.2f}"
            if not sens_v and chg_v:
                f_chg = _parse_first_float(chg_v)
                if f_chg is not None:
                    sens_v = f"{f_chg / i:.2f}"
            if amp_v: _set_cell_text(t0.rows[i].cells[1], amp_v, pt=10.0)
            if chg_v: _set_cell_text(t0.rows[i].cells[2], chg_v, pt=10.0)
            if sens_v:
                _set_cell_text(t0.rows[i].cells[3], sens_v, pt=10.0)
                f_s = _parse_first_float(sens_v)
                if f_s is not None: lin_sens_vals.append(f_s)

    freq_sens_vals = []
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for i in range(1, min(11, len(t1.rows))):
            amp_v = _gval(data, field_map, f"spec_{i}", f"freq_amp_{i}")
            chg_v = _gval(data, field_map, f"freq_charge_{i}")
            sens_v = _gval(data, field_map, f"act_{i}", f"freq_sens_{i}")
            f_amp = _parse_first_float(amp_v)
            if not chg_v and f_amp is not None:
                chg_v = f"{f_amp / 1.022:.2f}"
            if amp_v: _set_cell_text(t1.rows[i].cells[2], amp_v, pt=10.0)
            if chg_v: _set_cell_text(t1.rows[i].cells[3], chg_v, pt=10.0)
            if sens_v:
                _set_cell_text(t1.rows[i].cells[4], sens_v, pt=10.0)
                f_s = _parse_first_float(sens_v)
                if f_s is not None: freq_sens_vals.append(f_s)

    post_sens_vals = []
    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for i in range(1, min(11, len(t2.rows))):
            amp_v = _gval(data, field_map, f"post_amp_{i}")
            chg_v = _gval(data, field_map, f"post_charge_{i}")
            sens_v = _gval(data, field_map, f"post_sens_{i}")
            if amp_v: _set_cell_text(t2.rows[i].cells[1], amp_v, pt=10.0)
            if chg_v: _set_cell_text(t2.rows[i].cells[2], chg_v, pt=10.0)
            if sens_v:
                _set_cell_text(t2.rows[i].cells[3], sens_v, pt=10.0)
                f_s = _parse_first_float(sens_v)
                if f_s is not None: post_sens_vals.append(f_s)

    mean_lin = _gval(data, field_map, "mean_lin") or (f"{sum(lin_sens_vals)/len(lin_sens_vals):.2f}" if lin_sens_vals else "")
    mean_freq = _gval(data, field_map, "mean_freq") or (f"{sum(freq_sens_vals)/len(freq_sens_vals):.2f}" if freq_sens_vals else "")
    mean_post = _gval(data, field_map, "mean_post") or (f"{sum(post_sens_vals)/len(post_sens_vals):.2f}" if post_sens_vals else "")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("TEST REPORT FOR ACCELEROMETER:", lambda p, _: _set_para_text(p, f"TEST REPORT FOR ACCELEROMETER: MODEL NO {mod}", pt=11.0, bold=True) if mod else None),
        ("Job No:", lambda p, _: _set_para_text(p, f"Job No: {job}", pt=10.5)),
        ("Customer Name:", lambda p, _: _set_para_text(p, f"Customer Name: {cust}", pt=10.5)),
        ("Serial No:", lambda p, _: _set_para_text(p, f"Serial No: {ser}", pt=10.5)),
        ("Mean:", lambda p, _: _set_para_text(p, f"Mean: {mean_lin} pC/g", pt=10.5)),
        ("Mean:", lambda p, _: _set_para_text(p, f"Mean: {mean_freq} pC/g", pt=10.5)),
        ("Mean:", lambda p, _: _set_para_text(p, f"Mean: {mean_post} pC/g", pt=10.5)),
    ])

    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        for r_idx, key in enumerate(("charge_sens", "capacitance", "insulation_res")):
            v = _gval(data, field_map, key)
            if r_idx < len(t3.rows) and v:
                _set_cell_text(t3.rows[r_idx].cells[1], f": {v}", pt=10.0)

    if len(doc.tables) > 4:
        t4 = doc.tables[4]
        tb = _gval(data, field_map, "tested_by", "user_name")
        ab = _gval(data, field_map, "approved_by")
        dt = _gval(data, field_map, "date")
        if len(t4.rows) > 0:
            _set_cell_text(t4.rows[0].cells[1], f": {tb}", pt=10.0)
            _set_cell_text(t4.rows[0].cells[2], f"APPROVED BY: {ab}", pt=10.0)
        if len(t4.rows) > 1:
            _set_cell_text(t4.rows[1].cells[1], f": {dt}", pt=10.0)


def _fill_vibration_meter(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    vm_m = _gval(data, field_map, "vm_model", "model_no")
    vm_s = _gval(data, field_map, "vm_serial", "serial_no")
    acc_m = _gval(data, field_map, "acc_model")
    acc_s = _gval(data, field_map, "acc_serial")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("TESTED BY:", lambda p, _: set_para_with_tab(p, f"TESTED BY: {tb}", f"APPROVED BY: {ab}", 6200, 10.5, bold=False)),
        ("DATE:", lambda p, _: set_para_with_tab(p, f"DATE: {dt}", "(HOD-PDN)", 6200, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        t0 = doc.tables[0]
        _fill_cell_ordered_labels(t0.rows[0].cells[0], [
            ("JOB ORDER", job, None),
            ("CUSTOMER", cust, None),
        ], pt=10.0)
        _fill_cell_ordered_labels(t0.rows[1].cells[0], [
            ("MODEL NO.", vm_m, None),
            ("SERIAL NO.", vm_s, None),
        ], pt=10.0)
        _fill_cell_ordered_labels(t0.rows[1].cells[1], [
            ("MODEL NO.", acc_m, None),
            ("SERIAL NO.", acc_s, None),
        ], pt=10.0)

    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        # Rows 2..24 hold up to 23 measurement points; cols 4 & 5 are horizontally merged so unique cells are 6 columns
        for r_idx in range(2, len(t1.rows)):
            step = r_idx - 1
            uniq = _unique_row_cells(t1.rows[r_idx])
            if len(uniq) < 6:
                continue
            keys = (
                f"acc_std_{step}", f"acc_test_{step}",
                f"vel_std_{step}", f"vel_test_{step}",
                f"disp_std_{step}", f"disp_test_{step}",
            )
            for c_cell, k in zip(uniq, keys):
                v = _gval(data, field_map, k)
                if v: _set_cell_text(c_cell, v, pt=9.5)


def _fill_charge_amplifier(doc, data, field_map):
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    mod = _gval(data, field_map, "model_no")
    ser = _gval(data, field_map, "serial_no")
    acc_m = _gval(data, field_map, "acc_model")
    acc_s = _gval(data, field_map, "acc_serial")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("TESTED BY:", lambda p, _: set_para_with_tab(p, f"TESTED BY: {tb}", f"APPROVED BY: {ab}", 6200, 10.5, bold=False)),
        ("DATE:", lambda p, _: set_para_with_tab(p, f"DATE: {dt}", "HOD (PDN)", 6200, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        if len(t0.rows) > 0:
            _fill_cell_ordered_labels(t0.rows[0].cells[0], [
                ("JOB ORDER", job, None),
                ("CUSTOMER", cust, None),
            ], pt=10.0)
        if len(t0.rows) > 1:
            _fill_cell_ordered_labels(t0.rows[1].cells[0], [
                ("MODEL NO.", mod, None),
                ("SERIAL NO.", ser, None),
            ], pt=10.0)
            _fill_cell_ordered_labels(t0.rows[1].cells[1], [
                ("MODEL NO.", acc_m, None),
                ("SERIAL NO.", acc_s, None),
            ], pt=10.0)
        if len(t0.rows) > 2:
            spec_lines = [
                "",
                f": {_gval(data, field_map, 'freq_range')}",
                f": {_gval(data, field_map, 'max_in')}",
                f": {_gval(data, field_map, 'max_out')}",
                f": {_gval(data, field_map, 'conversion_factor')}",
                f": {_gval(data, field_map, 'pct_error')}",
            ]
            _set_cell_text(t0.rows[2].cells[1], "\n".join(spec_lines), pt=9.5)

    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for i in range(1, min(7, len(t1.rows))):
            in_v = _gval(data, field_map, f"spec_{i}", f"cal_in_{i}")
            out_v = _gval(data, field_map, f"act_{i}", f"cal_out_{i}")
            conv_v = _gval(data, field_map, f"conv_{i}")
            if not conv_v and in_v and out_v:
                f_in = _parse_first_float(in_v)
                f_out = _parse_first_float(out_v)
                if f_in and f_out is not None:
                    conv_v = f"{f_out / f_in:.2f}"
            if in_v: _set_cell_text(t1.rows[i].cells[1], in_v, pt=10.0)
            if out_v: _set_cell_text(t1.rows[i].cells[2], out_v, pt=10.0)
            if conv_v: _set_cell_text(t1.rows[i].cells[3], conv_v, pt=10.0)

    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for i in range(1, min(6, len(t2.rows))):
            vals = (
                _gval(data, field_map, f"freq_{i}"),
                _gval(data, field_map, f"freq_in_{i}"),
                _gval(data, field_map, f"freq_out_{i}"),
                _gval(data, field_map, f"freq_conv_{i}"),
            )
            for c_idx, v in enumerate(vals, start=1):
                if v: _set_cell_text(t2.rows[i].cells[c_idx], v, pt=10.0)


def _fill_odd_or_dd_system(doc, data, field_map, is_odd=True):
    if not doc.tables:
        return
    outer = doc.tables[0].cell(0, 0)
    job = _gval(data, field_map, "job_no", "job")
    cust = _gval(data, field_map, "customer", "cust")
    t1_out = _gval(data, field_map, "odd_trip_output" if is_odd else "dd1_trip_output", "trip_out_1")
    t2_out = _gval(data, field_map, "ddd_trip_output" if is_odd else "dd2_trip_output", "trip_out_2")
    lc_fail = _gval(data, field_map, "lc_fail_card", "lc_fail_err")
    ps_fail = _gval(data, field_map, "ps_fail_card", "ps_fail_err")
    rem = _gval(data, field_map, "remarks")

    _fill_paras_by_prefix(outer.paragraphs, [
        ("Job No", lambda p, _: _set_para_text(p, f"Job No.: {job}", pt=10.0)),
        ("Customer", lambda p, _: _set_para_text(p, f"Customer: {cust}", pt=10.0)),
        (("ODDD Trip Output", "DD1 Trip Output"), lambda p, s: _set_para_text(p, f"{s.split(':')[0]}: {t1_out}", pt=10.0)),
        (("DDD Trip Output", "DD2 Trip Output"), lambda p, s: _set_para_text(p, f"{s.split(':')[0]}: {t2_out}", pt=10.0)),
        (("Load Cell Failure", "Loadcell Failure"), lambda p, s: _set_para_text(p, f"{s.split(':')[0]}: {lc_fail}", pt=10.0)),
        ("Power Supply Failure", lambda p, s: _set_para_text(p, f"{s.split(':')[0]}: {ps_fail}", pt=10.0)),
        ("Remarks", lambda p, _: _set_para_text(p, f"Remarks (if any): {rem}", pt=10.0) if rem else None),
    ])

    nts = outer.tables
    if len(nts) > 0:
        nt0 = nts[0]
        if len(nt0.rows) > 4:
            for r_i, k1, k2 in (
                (1, "ind_model", "beam_serial_1"),
                (2, "ind_serial", "beam_serial_2"),
                (3, "ind_batch", "beam_batch"),
                (4, "ind_part", "beam_part"),
            ):
                v1, v2 = _gval(data, field_map, k1), _gval(data, field_map, k2)
                if v1: _set_cell_text(nt0.rows[r_i].cells[1], f": {v1}", pt=9.5)
                if v2: _set_cell_text(nt0.rows[r_i].cells[3], f": {v2}", pt=9.5)
        if len(nt0.rows) > 10:
            for r_i, k1, k2 in (
                (7, "lcl_model", "lcr_model"),
                (8, "lcl_serial", "lcr_serial"),
                (9, "lcl_batch", "lcr_batch"),
                (10, "lcl_part", "lcr_part"),
            ):
                v1, v2 = _gval(data, field_map, k1), _gval(data, field_map, k2)
                if v1: _set_cell_text(nt0.rows[r_i].cells[1], f": {v1}", pt=9.5)
                if v2: _set_cell_text(nt0.rows[r_i].cells[3], f": {v2}", pt=9.5)

    def _fill_trial_tbl(nt, pref, has_center=False):
        for t_i in range(1, 4):
            r_idx = t_i + 1
            if r_idx >= len(nt.rows):
                continue
            row = nt.rows[r_idx]
            ld = _gval(data, field_map, f"{pref}_load_{t_i}")
            lf = _gval(data, field_map, f"{pref}_left_{t_i}")
            if ld: _set_cell_text(row.cells[1], ld, pt=9.5)
            if lf: _set_cell_text(row.cells[2], lf, pt=9.5)
            if has_center:
                ct = _gval(data, field_map, f"{pref}_center_{t_i}")
                rt = _gval(data, field_map, f"{pref}_right_{t_i}")
                if ct: _set_cell_text(row.cells[3], ct, pt=9.5)
                if rt and len(row.cells) > 4: _set_cell_text(row.cells[4], rt, pt=9.5)
            else:
                rt = _gval(data, field_map, f"{pref}_right_{t_i}")
                if rt and len(row.cells) > 3: _set_cell_text(row.cells[3], rt, pt=9.5)

    if len(nts) > 1:
        _fill_trial_tbl(nts[1], "t1", has_center=is_odd and len(nts[1].columns) >= 5)
    if len(nts) > 2:
        _fill_trial_tbl(nts[2], "t2", has_center=False)
    if len(nts) > 3:
        nt3 = nts[3]
        tb = _gval(data, field_map, "tested_by", "user_name")
        ab = _gval(data, field_map, "approved_by")
        dt = _gval(data, field_map, "date")
        if len(nt3.rows) > 0:
            if tb: _set_cell_text(nt3.rows[0].cells[1], tb, pt=9.5)
            if ab: _set_cell_text(nt3.rows[0].cells[2], f"Approved by: {ab}", pt=9.5)
        if len(nt3.rows) > 1 and dt:
            _set_cell_text(nt3.rows[1].cells[1], dt, pt=9.5)


def _fill_performance_index(doc, data, field_map):
    tb = _gval(data, field_map, "prepared_by", "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    _fill_paras_by_prefix(doc.paragraphs, [
        ("Prepared", lambda p, _: _set_para_text(p, f"\t\t\t\tPrepared by\t: {tb}", pt=10.5)),
        ("(HOD", lambda p, _: _set_para_text(p, f"\t\t\t\t{ab} (HOD PDN)" if ab else "\t\t\t\t(HOD PDN)", pt=10.5)),
    ])
    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        month = _gval(data, field_map, "month", "month_year")
        sched = _gval(data, field_map, "scheduled", "jobs_scheduled")
        comp = _gval(data, field_map, "completed", "jobs_completed")
        pct = _gval(data, field_map, "pct_achieved")
        if not pct and sched and comp:
            fs, fc = _parse_first_float(sched), _parse_first_float(comp)
            if fs and fc is not None:
                pct = f"{(fc / fs) * 100.0:.2f}%"
        r1 = doc.tables[0].rows[1]
        for idx, v in enumerate((month, sched, comp, pct)):
            if v: _set_cell_text(r1.cells[idx], v, pt=10.5)


def _fill_performance_delay(doc, data, field_map):
    my = _gval(data, field_map, "month_year", "month")
    ts = _gval(data, field_map, "total_scheduled")
    te = _gval(data, field_map, "total_executed")
    td = _gval(data, field_map, "total_delay")
    pb = _gval(data, field_map, "prepared_by", "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("MONTH / YEAR:", lambda p, _: _set_para_text(p, f"MONTH / YEAR: {my}", pt=10.5)),
        ("TOTAL NO OF JOBS SCHEDULED", lambda p, _: _set_para_text(p, f"\tTOTAL NO OF JOBS SCHEDULED : {ts} Jobs", pt=10.5)),
        ("NO. OF JOBS EXECUTED", lambda p, _: _set_para_text(p, f"\tNO. OF JOBS EXECUTED        : {te} Jobs", pt=10.5)),
        ("TOTAL DELAY", lambda p, _: _set_para_text(p, f"\tTOTAL DELAY                                     : {td}", pt=10.5)),
        ("PREPARED BY:", lambda p, _: set_para_with_tab(p, f"\tPREPARED BY: {pb}", f"APPROVED BY: {ab}", 5500, 10.5, bold=False)),
        ("DATE:", lambda p, _: set_para_with_tab(p, f"              DATE: {dt}", "(HOD PDN)", 5500, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        rows_data = data.get("delay_rows")
        col_keys = ("month", "sched", "comp", "rework", "manpower", "shortage", "testing", "design")
        for i in range(1, 13):
            r_idx = i + 1
            if r_idx >= len(t0.rows) - 1:
                break
            row_vals = None
            if isinstance(rows_data, list) and (i - 1) < len(rows_data) and isinstance(rows_data[i - 1], (list, dict)):
                item = rows_data[i - 1]
                if isinstance(item, dict):
                    row_vals = [str(item.get(k) or "") for k in col_keys]
                else:
                    row_vals = [str(x or "") for x in item[:8]]
            if row_vals is None:
                row_vals = [_gval(data, field_map, f"{k}_{i}") for k in col_keys]
            for c_idx, v in enumerate(row_vals):
                if c_idx < len(t0.rows[r_idx].cells) and v.strip():
                    _set_cell_text(t0.rows[r_idx].cells[c_idx], v.strip(), pt=9.5)
        if len(t0.rows) > 14 and pb:
            _set_cell_text(t0.rows[14].cells[0], f"PREPARED BY: {pb}", pt=9.5)

    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        details = data.get("delay_items")
        if not isinstance(details, list) or not details:
            one = {
                "sl_no": _gval(data, field_map, "sl_no"),
                "jo_no": _gval(data, field_map, "jo_no", "job_no"),
                "customer": _gval(data, field_map, "customer"),
                "department": _gval(data, field_map, "department"),
                "reason": _gval(data, field_map, "reason"),
            }
            details = [one] if any(one.values()) else []
        for idx, item in enumerate(details, start=1):
            if idx >= len(t1.rows):
                t1.add_row()
            if isinstance(item, dict):
                vals = (item.get("sl_no", ""), item.get("jo_no", ""), item.get("customer", ""), item.get("department", ""), item.get("reason", ""))
            else:
                vals = list(item)[:5]
            for c_idx, v in enumerate(vals):
                if v: _set_cell_text(t1.rows[idx].cells[c_idx], str(v), pt=9.5)


def _fill_equipment_list(doc, data, field_map):
    if not doc.tables:
        return
    t0 = doc.tables[0]
    items = data.get("items")
    keys = ("sl", "desc", "ref", "cal_date", "due_date", "agency", "remarks")
    if isinstance(items, list) and items:
        for idx, item in enumerate(items, start=1):
            if idx >= len(t0.rows):
                t0.add_row()
            vals = ([item.get(k, "") for k in keys] if isinstance(item, dict) else list(item)[:7])
            for c_idx, v in enumerate(vals):
                if v: _set_cell_text(t0.rows[idx].cells[c_idx], str(v), pt=9.5)
        return

    for i in range(1, max(21, len(t0.rows))):
        vals = [_gval(data, field_map, f"eq_{i}_{k}") for k in keys]
        if any(vals):
            if i >= len(t0.rows):
                t0.add_row()
            for c_idx, v in enumerate(vals):
                if v: _set_cell_text(t0.rows[i].cells[c_idx], v, pt=9.5)


def _fill_loss_in_weigh_feeder(doc, data, field_map):
    rl = _gval(data, field_map, "res_load", "res_l")
    rr = _gval(data, field_map, "res_rate", "res_r")
    rt = _gval(data, field_map, "res_tot", "res_t")
    span_r = _gval(data, field_map, "span_rate")
    pv = _gval(data, field_map, "p_val")
    iv = _gval(data, field_map, "i_val")
    dv = _gval(data, field_map, "d_val")
    ff = _gval(data, field_map, "feed_fwd_time")
    ri = _gval(data, field_map, "rate_interval_time")
    rem = _gval(data, field_map, "remarks")
    inst = _gval(data, field_map, "instrument_used")
    tb = _gval(data, field_map, "tested_by", "user_name")
    ab = _gval(data, field_map, "approved_by")
    dt = _gval(data, field_map, "date")

    _fill_paras_by_prefix(doc.paragraphs, [
        ("Resolution:", lambda p, _: _set_para_text(p, f"Resolution: Load: {rl} kg/m   Rate: {rr} tph   Totalizer: {rt} tonnes", pt=10.0)),
        ("Span Rate:", lambda p, _: _set_para_text(p, f"Span Rate: {span_r} tph / kg/hr", pt=10.0)),
        ("P Value:", lambda p, _: _set_para_text(p, f"P Value: {pv}        I Value: {iv}        D Value: {dv}", pt=10.0)),
        ("Feed Forward time:", lambda p, _: set_para_with_tab(p, f"Feed Forward time: {ff}", f"Rate Interval Time: {ri}", 4800, 10.0, bold=False)),
        ("Remarks:", lambda p, _: _set_para_text(p, f"Remarks: {rem}", pt=10.0)),
        ("Instrument Used:", lambda p, _: _set_para_text(p, f" Instrument Used: {inst}", pt=10.0)),
        ("Tested by:", lambda p, _: set_para_with_tab(p, f"\tTested by: {tb}", f"Approved by: {ab}", 6200, 10.5, bold=False)),
        ("Date:", lambda p, _: set_para_with_tab(p, f"\tDate: {dt}", "(HOD – PDN)", 6200, 10.5, bold=False)),
    ])

    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        r1 = doc.tables[0].rows[1]
        _fill_cell_ordered_labels(r1.cells[0], [
            ("Model No.", _gval(data, field_map, "liw_model", "model_no"), None),
            ("Serial No.", _gval(data, field_map, "liw_serial", "serial_no"), None),
            ("Model No.", _gval(data, field_map, "lp_model"), None),
            ("Serial No.", _gval(data, field_map, "lp_serial"), None),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[1], [
            ("Model No.", _gval(data, field_map, "ctrl_model"), None),
            ("Serial No.", _gval(data, field_map, "ctrl_serial"), None),
            ("SW version", _gval(data, field_map, "ctrl_sw", "sw_version"), None),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[2], [
            ("Model No.", _gval(data, field_map, "lc_model"), None),
            ("1)", _gval(data, field_map, "lc_serial_1", "s1"), lambda _, v: f"1) {v}"),
            ("2)", _gval(data, field_map, "lc_serial_2", "s2"), lambda _, v: f"2) {v}"),
            ("3)", _gval(data, field_map, "lc_serial_3", "s3"), lambda _, v: f"3) {v}"),
            ("4)", _gval(data, field_map, "lc_serial_4", "s4"), lambda _, v: f"4) {v}"),
        ], pt=9.0)
        _fill_cell_ordered_labels(r1.cells[3], [
            ("Model No.", _gval(data, field_map, "jbox_1_model", "jbox_model1"), None),
            ("Serial No.", _gval(data, field_map, "jbox_1_serial", "jbox_serial1"), None),
            ("Model No.", _gval(data, field_map, "jbox_2_model", "jbox_model2"), None),
            ("Serial No.", _gval(data, field_map, "jbox_2_serial", "jbox_serial2"), None),
            ("Model No.", _gval(data, field_map, "jbox_3_model", "jbox_model3"), None),
            ("Serial No.", _gval(data, field_map, "jbox_3_serial", "jbox_serial3"), None),
        ], pt=9.0)

    if len(doc.tables) > 1 and len(doc.tables[1].rows) > 2:
        r2 = doc.tables[1].rows[2]
        _fill_cell_ordered_labels(r2.cells[0], [
            ("Make", _gval(data, field_map, "drive_make"), None),
            ("Model", _gval(data, field_map, "drive_model"), None),
            ("Serial", _gval(data, field_map, "drive_serial"), None),
            ("Capacity", _gval(data, field_map, "drive_kw"), lambda _, v: f"Capacity: {v} kW" if v else "Capacity: kW"),
        ], pt=9.0)
        _fill_cell_ordered_labels(r2.cells[1], [
            ("Make", _gval(data, field_map, "screw_motor_make"), None),
            ("Model No.", _gval(data, field_map, "screw_motor_model"), None),
            ("Serial", _gval(data, field_map, "screw_motor_serial"), None),
            ("Rating", _gval(data, field_map, "screw_motor_kw"), lambda _, v: f"Rating: {v} kW" if v else "Rating: kW"),
        ], pt=9.0)
        _fill_cell_ordered_labels(r2.cells[2], [
            ("Make", _gval(data, field_map, "screw_gb_make"), None),
            ("Model No.", _gval(data, field_map, "screw_gb_model"), None),
            ("Serial", _gval(data, field_map, "screw_gb_serial"), None),
            ("Ratio", _gval(data, field_map, "screw_gb_ratio"), None),
        ], pt=9.0)
        _fill_cell_ordered_labels(r2.cells[3], [
            ("Make", _gval(data, field_map, "agit_motor_make"), None),
            ("Model No.", _gval(data, field_map, "agit_motor_model"), None),
            ("Serial", _gval(data, field_map, "agit_motor_serial"), None),
            ("Rating", _gval(data, field_map, "agit_motor_kw"), lambda _, v: f"Rating: {v} kW" if v else "Rating: kW"),
        ], pt=9.0)
        _fill_cell_ordered_labels(r2.cells[4], [
            ("Make", _gval(data, field_map, "agit_gb_make"), None),
            ("Model No.", _gval(data, field_map, "agit_gb_model"), None),
            ("Serial", _gval(data, field_map, "agit_gb_serial"), None),
            ("Ratio", _gval(data, field_map, "agit_gb_ratio"), None),
        ], pt=9.0)

    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for i in range(1, min(6, len(t2.rows))):
            sp = _gval(data, field_map, f"spec_{i}")
            ac = _gval(data, field_map, f"act_{i}")
            if sp: _set_cell_text(t2.rows[i].cells[2], sp, pt=9.5)
            if ac: _set_cell_text(t2.rows[i].cells[3], ac, pt=9.5)

    if len(doc.tables) > 3:
        t3 = doc.tables[3]
        for i in range(1, min(7, len(t3.rows))):
            vals = (
                _gval(data, field_map, f"set_rate_{i}"),
                _gval(data, field_map, f"achieved_rate_{i}"),
                _gval(data, field_map, f"pct_err_{i}"),
                _gval(data, field_map, f"ma_out_{i}"),
            )
            for c_idx, v in enumerate(vals):
                if v: _set_cell_text(t3.rows[i].cells[c_idx], v, pt=9.5)

    if len(doc.tables) > 4:
        t4 = doc.tables[4]
        for r_i, fk in enumerate(("func_mode", "func_di", "func_pf", "func_aio", "func_comm")):
            v = _gval(data, field_map, fk)
            if v and r_i < len(t4.rows):
                _set_cell_text(t4.rows[r_i].cells[1], v, pt=9.5)


def _apply_template_specific_fillers(doc, report_type, data, field_map):
    if report_type == "inprocess":
        st = _gval(data, field_map, "inspection_status")
        if st and doc.tables and len(doc.tables[0].rows) > 5:
            _set_cell_text(doc.tables[0].rows[5].cells[1], st, pt=10.5)
        return
    if report_type == "work_instructions":
        st = _gval(data, field_map, "compliance_status")
        if st and doc.tables and len(doc.tables[0].rows) > 5:
            _set_cell_text(doc.tables[0].rows[5].cells[1], st, pt=10.5)
        return

    handlers = {
        "job_traveller_card": _fill_job_traveller_card,
        "batching_system": _fill_batching_system,
        "remote_indicator": _fill_remote_indicator,
        "weigh_feeder": _fill_weigh_feeder,
        "trip_safe": _fill_trip_safe,
        "vibration_switch": _fill_vibration_switch,
        "acc_mv": _fill_acc_mv,
        "acc_charge": _fill_acc_charge,
        "vibration_meter": _fill_vibration_meter,
        "charge_amplifier": _fill_charge_amplifier,
        "odd_system": lambda d, dt, fm: _fill_odd_or_dd_system(d, dt, fm, is_odd=True),
        "dd_system": lambda d, dt, fm: _fill_odd_or_dd_system(d, dt, fm, is_odd=False),
        "performance_index": _fill_performance_index,
        "performance_delay": _fill_performance_delay,
        "equipment_list": _fill_equipment_list,
        "loss_in_weigh_feeder": _fill_loss_in_weigh_feeder,
    }
    handler = handlers.get(report_type)
    if handler:
        handler(doc, data, field_map)


def generate_docx_record(data, output_path):
    report_type = str(data.get("report_type") or "belt_scale").strip()
    if report_type == "inprocess_register":
        report_type = "inprocess"
    elif report_type == "performance_delay_analysis":
        report_type = "performance_delay"
    if report_type in ["crane_scale", "cws"]:
        return generate_crane_scale_record(data, output_path)
    if report_type == "digital_indicator":
        return generate_digital_indicator_record(data, output_path)
    if report_type == "signal_conditioner":
        return generate_signal_conditioner_record(data, output_path)
    if report_type in ["misc_report", "misc", "miscellaneous", "miscellaneous_items"]:
        return generate_misc_record(data, output_path)

    template_name = TEMPLATE_MAP.get(report_type)
    if report_type == "belt_scale":
        raw_ao = data.get("num_outputs")
        if raw_ao is not None and str(raw_ao).strip() in ["1", "2", "3"]:
            num_ao = str(raw_ao).strip()
        else:
            has_ao3 = any(str(data.get(f"g{r}_5_ao3") or "").strip() for r in range(1, 12))
            has_ao2 = any(str(data.get(f"g{r}_5_ao2") or "").strip() for r in range(1, 12))
            if has_ao3: num_ao = "3"
            elif has_ao2: num_ao = "2"
            else: num_ao = "1"

        lc_system = str(data.get("lc_system") or "").strip()
        if not lc_system:
            has_8lc = any(str(data.get(f"s{i}") or data.get(f"lc_s{i}") or "").strip() for i in range(5, 9))
            lc_system = "8" if has_8lc else "4"

        if lc_system == "8":
            if num_ao == "3":
                template_name = "PP-05_03 Belt Scale - 3 AO - 8 LC.docx"
            elif num_ao == "2":
                template_name = "PP-05_03 Belt Scale - 2 AO - 8 LC.docx"
            else:
                template_name = "PP-05_03 Belt Scale - 8 LC.docx"
        else:
            if num_ao == "3":
                template_name = "PP-05_03 Belt Scale - 3 AO.docx"
            elif num_ao == "2":
                template_name = "PP-05_03 Belt Scale - 2 AO.docx"
            else:
                template_name = "PP-05_03 Belt Scale.docx"

    if not template_name:
        raise ValueError(f"Unknown report type: {report_type}")

    template_path = os.path.join(TEMPLATES_DIR, template_name)
    if not os.path.exists(template_path):
        doc_dir_path = os.path.join(BASE_DIR, "DOC", template_name)
        if os.path.exists(doc_dir_path):
            template_path = doc_dir_path
        else:
            raise FileNotFoundError(f"Template not found: {template_path} (or {doc_dir_path})")

    doc = docx.Document(template_path)

    # Convert all None to empty strings to avoid 'None' printing
    field_map = {k: (v if v is not None else "") for k, v in data.items()}

    # Flatten routine_tests array into spec_1..9 and act_1..9 (only if non-empty)
    if isinstance(data.get("routine_tests"), list):
        for idx, item in enumerate(data["routine_tests"], start=1):
            if isinstance(item, dict):
                spec_v = item.get("spec")
                act_v = item.get("act") or item.get("actual")
                if spec_v is not None and str(spec_v).strip():
                    field_map[f"spec_{idx}"] = str(spec_v)
                if act_v is not None and str(act_v).strip():
                    field_map[f"act_{idx}"] = str(act_v)

    # Flatten grid_data array into g1_0..g11_5 (only if non-empty)
    if isinstance(data.get("grid_data"), list):
        for r_idx, row in enumerate(data["grid_data"], start=1):
            if isinstance(row, list):
                for c_idx, val in enumerate(row):
                    str_val = str(val) if val is not None else ""
                    if str_val.strip():
                        if report_type == 'belt_scale' and c_idx == 2:
                            g_speed = _parse_first_float(str_val)
                            if g_speed is not None:
                                str_val = f"{g_speed:.2f}"
                        field_map[f"g{r_idx}_{c_idx}"] = str_val
                        if c_idx == 5:
                            field_map[f"g{r_idx}_5"] = str_val
                            field_map[f"g{r_idx}5"] = str_val
                            field_map[f"g{r_idx}_5_ao1"] = str_val
                            field_map[f"g{r_idx}5_ao1"] = str_val
                        elif c_idx == 6:
                            field_map[f"g{r_idx}_6"] = str_val
                            field_map[f"g{r_idx}6"] = str_val
                            field_map[f"g{r_idx}_5_ao2"] = str_val
                            field_map[f"g{r_idx}5_ao2"] = str_val
                        elif c_idx == 7:
                            field_map[f"g{r_idx}_7"] = str_val
                            field_map[f"g{r_idx}7"] = str_val
                            field_map[f"g{r_idx}_5_ao3"] = str_val
                            field_map[f"g{r_idx}5_ao3"] = str_val

    # Ensure aliasing for Column 5 if g{r}_5_ao1 was entered
    for r in range(1, 12):
        if not field_map.get(f"g{r}_5") and field_map.get(f"g{r}_5_ao1"):
            field_map[f"g{r}_5"] = field_map[f"g{r}_5_ao1"]
            field_map[f"g{r}5"] = field_map[f"g{r}_5_ao1"]

    # System-specific aliases & default fallbacks
    if report_type == 'belt_scale':
        def get_first(*keys):
            for k in keys:
                v = str(field_map.get(k) or data.get(k) or '').strip()
                if v:
                    return v
            return ''

        cust_val = get_first('customer', 'customer_name')
        field_map['customer'] = cust_val
        field_map['customer_name'] = cust_val

        bs_m_val = get_first('bs_model', 'model_no', 'belt_scale_model')
        field_map['bs_model'] = bs_m_val
        field_map['model_no'] = bs_m_val
        field_map['belt_scale_model'] = bs_m_val

        bs_s_val = get_first('bs_serial', 'serial_no', 'belt_scale_serial')
        field_map['bs_serial'] = bs_s_val
        field_map['serial_no'] = bs_s_val
        field_map['belt_scale_serial'] = bs_s_val

        sw_val = get_first('bs_sw', 'version', 'sw_version')
        field_map['bs_sw'] = sw_val
        field_map['version'] = sw_val

        rm_m_val = get_first('rm_model', 'remote_model')
        field_map['rm_model'] = rm_m_val
        field_map['remote_model'] = rm_m_val

        rm_s_val = get_first('rm_serial', 'remote_serial')
        field_map['rm_serial'] = rm_s_val
        field_map['remote_serial'] = rm_s_val

        lc_m_val = get_first('lc_model', 'sensor_model', 'load_cell_model')
        field_map['lc_model'] = lc_m_val
        field_map['sensor_model'] = lc_m_val

        # Support lc_serials list if passed
        lc_serials_arr = data.get('lc_serials') or field_map.get('lc_serials')
        if isinstance(lc_serials_arr, list):
            for i_lc, s_val in enumerate(lc_serials_arr, start=1):
                if i_lc <= 8:
                    field_map[f's{i_lc}'] = str(s_val).strip()

        for i_s in range(1, 9):
            field_map[f's{i_s}'] = get_first(f's{i_s}', f'lc_s{i_s}', f'lc_serial_{i_s}', f'sensor_s{i_s}')

        # 8LC lines with clean column alignment preventing awkward text wrapping
        def format_8lc_line(num1, num2):
            val1 = str(field_map.get(f's{num1}') or '').strip()
            val2 = str(field_map.get(f's{num2}') or '').strip()
            left_part = f"{num1}) {val1}".strip() if val1 else f"{num1})"
            right_part = f"{num2}) {val2}".strip() if val2 else f"{num2})"
            return f"{left_part}\t{right_part}"

        field_map['lc_line_1'] = format_8lc_line(1, 5)
        field_map['lc_line_2'] = format_8lc_line(2, 6)
        field_map['lc_line_3'] = format_8lc_line(3, 7)
        field_map['lc_line_4'] = format_8lc_line(4, 8)

        ss_m_val = get_first('ss_model', 'tacho_model', 'speed_sensor_model')
        field_map['ss_model'] = ss_m_val
        field_map['tacho_model'] = ss_m_val

        ss_s_val = get_first('ss_serial', 'tacho_serial', 'speed_sensor_serial')
        field_map['ss_serial'] = ss_s_val
        field_map['tacho_serial'] = ss_s_val

        as_m_val = get_first('as_model', 'angle_sensor_model')
        field_map['as_model'] = as_m_val
        field_map['angle_sensor_model'] = as_m_val

        as_s_val = get_first('as_serial', 'angle_sensor_serial')
        field_map['as_serial'] = as_s_val
        field_map['angle_sensor_serial'] = as_s_val

        jb1_m_val = get_first('jb_1_model', 'jbox_model1')
        field_map['jb_1_model'] = jb1_m_val
        field_map['jbox_model1'] = jb1_m_val

        jb1_s_val = get_first('jb_1_serial', 'jbox_serial1')
        field_map['jb_1_serial'] = jb1_s_val
        field_map['jbox_serial1'] = jb1_s_val

        jb2_m_val = get_first('jb_2_model', 'jbox_model2')
        field_map['jb_2_model'] = jb2_m_val
        field_map['jbox_model2'] = jb2_m_val

        jb2_s_val = get_first('jb_2_serial', 'jbox_serial2')
        field_map['jb_2_serial'] = jb2_s_val
        field_map['jbox_serial2'] = jb2_s_val

        jb3_m_val = get_first('jb_3_model', 'jbox_model3')
        field_map['jb_3_model'] = jb3_m_val
        field_map['jbox_model3'] = jb3_m_val

        jb3_s_val = get_first('jb_3_serial', 'jbox_serial3')
        field_map['jb_3_serial'] = jb3_s_val
        field_map['jbox_serial3'] = jb3_s_val

        chk_val = get_first('test_voltage_chk', 'short_check')
        field_map['test_voltage_chk'] = chk_val
        field_map['short_check'] = chk_val
        
        # Ensure belt_speed fallback from tacho_belt_speed if missing
        if not field_map.get('belt_speed') and field_map.get('tacho_belt_speed'):
            field_map['belt_speed'] = field_map['tacho_belt_speed']

        # Format belt_speed to 2 decimal places if provided
        raw_speed_str = str(field_map.get('belt_speed') or '').strip()
        speed_val = _parse_first_float(raw_speed_str, 0.0) or 0.0
        if raw_speed_str and speed_val > 0 and 'm/s' not in raw_speed_str.lower():
            field_map['belt_speed'] = f"{speed_val:.2f} m/s"

        # Automatic Resolution calculation ONLY IF user entered capacity and speed
        raw_cap_str = str(field_map.get('capacity') or '').strip()
        cap_val_parsed = _parse_first_float(raw_cap_str, 0.0) or 0.0
        if raw_cap_str and cap_val_parsed > 0 and speed_val > 0:
            try:
                cap_val = cap_val_parsed
                full_load = (cap_val * 1000.0) / (3600.0 * speed_val)
                
                def calc_res(val):
                    v = abs(float(val))
                    if v <= 99: return '0.01'
                    if v <= 999: return '0.1'
                    return '1'

                if not field_map.get('res_l'): field_map['res_l'] = calc_res(full_load)
                if not field_map.get('res_r'): field_map['res_r'] = calc_res(cap_val)
                if not field_map.get('res_s'): field_map['res_s'] = '0.01'
                if not field_map.get('res_t'): field_map['res_t'] = '0.1'
            except Exception:
                pass

        # Determine decimal places for grid formatting
        def _get_res_dec(res_val, default=2):
            s = str(res_val or '').strip()
            if '.' in s:
                num_token = s.split()[0]
                if '.' in num_token:
                    return len(num_token.split('.')[1])
            if s.split()[0] == '1' if s else False:
                return 0
            return default

        dec_l = _get_res_dec(field_map.get('res_l') or field_map.get('rl'), 2)
        dec_r = _get_res_dec(field_map.get('res_r') or field_map.get('rr'), 2)
        dec_s = _get_res_dec(field_map.get('res_s') or field_map.get('rs'), 2)
        dec_t = _get_res_dec(field_map.get('res_t') or field_map.get('rt'), 1)
        bs_calib_mode = str(data.get("calib_mode") or "").strip().lower()

        # Re-format measurement grid values based on resolution
        if isinstance(data.get("grid_data"), list):
            for r_idx, row in enumerate(data["grid_data"], start=1):
                if isinstance(row, list):
                    for c_idx, val in enumerate(row):
                        str_val = str(val).strip() if val is not None else ""
                        if str_val:
                            num_v = _parse_first_float(str_val)
                            if num_v is not None:
                                if c_idx == 0:  # mV or Load (kg/m)
                                    dec_0 = 2 if (bs_calib_mode == "mv" or "mv" in str_val.lower()) else dec_l
                                    field_map[f"g{r_idx}_0"] = f"{num_v:.{dec_0}f}"
                                elif c_idx == 1:  # Belt Kg/m
                                    field_map[f"g{r_idx}_1"] = f"{num_v:.{dec_l}f}"
                                elif c_idx == 2:  # Speed m/s
                                    field_map[f"g{r_idx}_2"] = f"{num_v:.{dec_s}f}"
                                elif c_idx == 3:  # Rate tph
                                    field_map[f"g{r_idx}_3"] = f"{num_v:.{dec_r}f}"
                                elif c_idx == 4:  # Totalizer Tonnes
                                    field_map[f"g{r_idx}_4"] = f"{num_v:.{dec_t}f}"
                                elif c_idx == 5:  # AO1 mA
                                    field_map[f"g{r_idx}_5"] = f"{num_v:.2f}"
                                elif c_idx == 6:  # AO2 mA
                                    field_map[f"g{r_idx}_5_ao2"] = f"{num_v:.2f}"
                                elif c_idx == 7:  # AO3 mA
                                    field_map[f"g{r_idx}_5_ao3"] = f"{num_v:.2f}"

        # Short abbreviated keys to match 100% with backend template
        field_map['job'] = str(field_map.get('job_no') or data.get('job_no') or '')
        field_map['cust'] = str(field_map.get('customer') or data.get('customer') or '')
        field_map['cap'] = str(field_map.get('capacity') or data.get('capacity') or '')
        field_map['cn'] = str(field_map.get('conveyor_no') or data.get('conveyor_no') or '')

        field_map['bm'] = str(field_map.get('bs_model') or '')
        field_map['bs'] = str(field_map.get('bs_serial') or '')
        field_map['sw'] = str(field_map.get('bs_sw') or field_map.get('version') or '')

        field_map['rmm'] = str(field_map.get('rm_model') or field_map.get('remote_model') or '')
        field_map['rms'] = str(field_map.get('rm_serial') or field_map.get('remote_serial') or '')

        field_map['lm'] = str(field_map.get('sensor_model') or field_map.get('lc_model') or '')
        field_map['s1'] = str(field_map.get('s1') or '')
        field_map['s2'] = str(field_map.get('s2') or '')
        field_map['s3'] = str(field_map.get('s3') or '')
        field_map['s4'] = str(field_map.get('s4') or '')
        field_map['s5'] = str(field_map.get('s5') or '')
        field_map['s6'] = str(field_map.get('s6') or '')
        field_map['s7'] = str(field_map.get('s7') or '')
        field_map['s8'] = str(field_map.get('s8') or '')

        field_map['sm'] = str(field_map.get('tacho_model') or field_map.get('ss_model') or '')
        field_map['ss'] = str(field_map.get('tacho_serial') or field_map.get('ss_serial') or '')

        field_map['am'] = str(field_map.get('angle_sensor_model') or field_map.get('as_model') or '')
        field_map['as'] = str(field_map.get('angle_sensor_serial') or field_map.get('as_serial') or '')

        field_map['j1m'] = str(field_map.get('jbox_model1') or field_map.get('jb_1_model') or '')
        field_map['j1s'] = str(field_map.get('jbox_serial1') or field_map.get('jb_1_serial') or '')
        field_map['j2m'] = str(field_map.get('jbox_model2') or field_map.get('jb_2_model') or '')
        field_map['j2s'] = str(field_map.get('jbox_serial2') or field_map.get('jb_2_serial') or '')
        field_map['j3m'] = str(field_map.get('jbox_model3') or field_map.get('jb_3_model') or '')
        field_map['j3s'] = str(field_map.get('jbox_serial3') or field_map.get('jb_3_serial') or '')

        # Check if input supply is 24V
        raw_sup = str(field_map.get('spec_1') or field_map.get('act_1') or field_map.get('sp1') or field_map.get('ac1') or '').strip().upper()
        is_24v = '24' in raw_sup

        # Table 1: Pre-check handling
        if is_24v:
            field_map['sc'] = 'CHECKED AND FOUND OK'
            if len(doc.tables) > 1 and len(doc.tables[1].rows) > 0 and len(doc.tables[1].rows[0].cells) > 0:
                t1_cell = doc.tables[1].rows[0].cells[0]
                if len(t1_cell.paragraphs) >= 2:
                    t1_cell.paragraphs[0].text = "Check for 24V DC                                          :   CHECKED AND FOUND OK"
                    t1_cell.paragraphs[1].text = ""
                elif len(t1_cell.paragraphs) == 1:
                    t1_cell.paragraphs[0].text = "Check for 24V DC                                          :   CHECKED AND FOUND OK"
        else:
            sc_val = str(field_map.get('test_voltage_chk') or field_map.get('short_check') or '').strip()
            if not sc_val or sc_val.upper() in ['OK', 'CHECKED', 'PASSED']:
                field_map['sc'] = 'CHECKED AND FOUND OK'
            else:
                field_map['sc'] = sc_val

        raw_rl = str(field_map.get('res_l') or field_map.get('rl') or '').strip()
        raw_rs = str(field_map.get('res_s') or field_map.get('rs') or '').strip()
        raw_rr = str(field_map.get('res_r') or field_map.get('rr') or '').strip()
        raw_rt = str(field_map.get('res_t') or field_map.get('rt') or '').strip()

        # Resolve res_l cleanly before appending " kg/m" unit to field_map['rl']
        res_l_token = re.sub(r'(?i)\s*kg/m', '', raw_rl).strip()
        field_map['rl'] = f"{raw_rl} kg/m" if raw_rl and 'kg' not in raw_rl.lower() else (raw_rl or "0.01 kg/m")
        field_map['rs'] = f"{raw_rs} m/s" if raw_rs and 'm/s' not in raw_rs.lower() else (raw_rs or "0.01 m/s")
        field_map['rr'] = f"{raw_rr} tph" if raw_rr and 'tph' not in raw_rr.lower() else (raw_rr or "0.1 tph")
        field_map['rt'] = f"{raw_rt} tonnes" if raw_rt and 'tonne' not in raw_rt.lower() else (raw_rt or "0.1 tonnes")

        # 1. Input Supply Voltage (Row 1) - default 230V
        sp1_val = str(field_map.get('spec_1') or field_map.get('sp1') or '230V').strip()
        ac1_val = str(field_map.get('act_1') or field_map.get('ac1') or '230V').strip()
        field_map['sp1'] = sp1_val if sp1_val and sp1_val.upper() != 'NA' else '230V'
        field_map['ac1'] = ac1_val if ac1_val and ac1_val.upper() != 'NA' else '230V'

        # 2. Derived DC Voltages (Row 2) - default 24V
        sp2_val = str(field_map.get('spec_2') or field_map.get('sp2') or '24V').strip()
        ac2_val = str(field_map.get('act_2') or field_map.get('ac2') or '24V').strip()
        field_map['sp2'] = sp2_val if sp2_val and sp2_val.upper() != 'NA' else '24V'
        field_map['ac2'] = ac2_val if ac2_val and ac2_val.upper() != 'NA' else '24V'

        # 3. Display and Keypad Functionality (Row 3) - preserve user input if entered
        sp3_val = str(field_map.get('spec_3') or field_map.get('sp3') or '').strip()
        ac3_val = str(field_map.get('act_3') or field_map.get('ac3') or '').strip()
        field_map['sp3'] = sp3_val if sp3_val and sp3_val.upper() != 'NA' else 'OK'
        field_map['ac3'] = ac3_val if ac3_val and ac3_val.upper() != 'NA' else 'OK'

        # 4. Belt Load (Row 4) - default 5.00 mV if spec_4/sp4 not provided; ac4 is clean load value matching resolution
        sp4_val = str(field_map.get('spec_4') or field_map.get('sp4') or '').strip()
        field_map['sp4'] = sp4_val if sp4_val and sp4_val.upper() != 'NA' else '5.00 mV'
        c_val = _parse_first_float(field_map.get('capacity') or field_map.get('rated_capacity'), 0.0) or 0.0
        s_val = _parse_first_float(field_map.get('belt_speed'), 0.0) or 0.0
        calc_load = (c_val * 1000.0) / (3600.0 * s_val) if (c_val > 0 and s_val > 0) else None

        res_l = res_l_token or str(field_map.get('res_l') or '').strip()
        if res_l == '0.01':
            dec_l = 2
        elif res_l == '0.1':
            dec_l = 1
        elif res_l == '1':
            dec_l = 0
        else:
            if calc_load is not None:
                dec_l = 2 if calc_load <= 99 else (1 if calc_load <= 999 else 0)
            else:
                dec_l = 2

        raw_ac4 = str(field_map.get('act_4') or field_map.get('ac4') or '').strip()
        clean_ac4_num = _parse_first_float(raw_ac4)
        if clean_ac4_num is not None:
            field_map['ac4'] = f"{clean_ac4_num:.{dec_l}f}"
        elif raw_ac4:
            field_map['ac4'] = raw_ac4
        elif calc_load is not None:
            field_map['ac4'] = f"{calc_load:.{dec_l}f}"
        else:
            field_map['ac4'] = "0.00"

        # 5. Belt Speed (Row 5) - sp5 is Frequency Hz, ac5 is clean speed value
        p_val = _parse_first_float(field_map.get('pulses') or field_map.get('speed_pulses'), 12.0) or 12.0
        d_val = _parse_first_float(field_map.get('pulley_diameter') or field_map.get('tacho_wheel_dia'), 0.16) or 0.16
        if d_val > 5.0:
            d_val = d_val / 1000.0
        calc_hz = (s_val * p_val) / (math.pi * d_val) if (s_val > 0 and d_val > 0) else None

        raw_sp5 = str(field_map.get('spec_5') or field_map.get('sp5') or '').strip()
        clean_sp5_num = _parse_first_float(raw_sp5)
        if clean_sp5_num is not None:
            field_map['sp5'] = f"{clean_sp5_num:.2f} Hz"
        elif raw_sp5:
            field_map['sp5'] = raw_sp5
        elif calc_hz is not None:
            field_map['sp5'] = f"{calc_hz:.2f} Hz"
        else:
            field_map['sp5'] = "31.83 Hz"

        raw_ac5 = str(field_map.get('act_5') or field_map.get('ac5') or '').strip()
        clean_ac5_num = _parse_first_float(raw_ac5)
        if clean_ac5_num is not None:
            field_map['ac5'] = f"{clean_ac5_num:.2f}"
        elif raw_ac5:
            field_map['ac5'] = raw_ac5
        elif s_val > 0:
            field_map['ac5'] = f"{s_val:.2f}"
        else:
            field_map['ac5'] = "2.50"

        # 6. Number of PF Contacts (Row 6)
        sp6_val = str(field_map.get('spec_6') or field_map.get('sp6') or '').strip()
        ac6_val = str(field_map.get('act_6') or field_map.get('ac6') or '').strip()
        field_map['sp6'] = sp6_val if sp6_val and sp6_val.upper() != 'NA' else '4'
        field_map['ac6'] = ac6_val if ac6_val and ac6_val.upper() != 'NA' else '4'

        # 7. Communication Output (Row 7)
        sp7_val = str(field_map.get('spec_7') or field_map.get('sp7') or '').strip()
        ac7_val = str(field_map.get('act_7') or field_map.get('ac7') or '').strip()
        field_map['sp7'] = sp7_val if sp7_val and sp7_val.upper() != 'NA' else 'RS 485'
        field_map['ac7'] = ac7_val if ac7_val and ac7_val.upper() != 'NA' else 'RS 485'

        # 8. Analog Output (Row 8) label based on num_ao if not specified
        raw_ao = data.get("num_outputs") or field_map.get("num_outputs")
        if raw_ao is not None and str(raw_ao).strip() in ["1", "2", "3"]:
            num_ao = str(raw_ao).strip()
        else:
            has_ao3 = any(str(data.get(f"g{r}_5_ao3") or field_map.get(f"g{r}_5_ao3") or "").strip() for r in range(1, 12))
            has_ao2 = any(str(data.get(f"g{r}_5_ao2") or field_map.get(f"g{r}_5_ao2") or "").strip() for r in range(1, 12))
            if has_ao3: num_ao = "3"
            elif has_ao2: num_ao = "2"
            else: num_ao = "1"

        default_ao_lbl = '3 x 4-20mA' if num_ao == '3' else ('2 x 4-20mA' if num_ao == '2' else '4-20mA')
        sp8_val = str(field_map.get('spec_8') or field_map.get('sp8') or '').strip()
        ac8_val = str(field_map.get('act_8') or field_map.get('ac8') or '').strip()
        field_map['sp8'] = sp8_val if sp8_val and sp8_val.upper() != 'NA' else default_ao_lbl
        field_map['ac8'] = ac8_val if ac8_val and ac8_val.upper() != 'NA' else default_ao_lbl

        # 9. Wiring, TB and Component Layout (Row 9)
        sp9_val = str(field_map.get('spec_9') or field_map.get('sp9') or '').strip()
        ac9_val = str(field_map.get('act_9') or field_map.get('ac9') or '').strip()
        field_map['sp9'] = sp9_val if sp9_val and sp9_val.upper() != 'NA' else 'OK'
        field_map['ac9'] = ac9_val if ac9_val and ac9_val.upper() != 'NA' else 'OK'

        # Map grid data g10..g115 and AO1, AO2, AO3 - DO NOT duplicate AO1 into AO2/AO3!
        for r_g in range(1, 12):
            for c_g in range(5):
                val_g = str(field_map.get(f'g{r_g}_{c_g}') or '')
                field_map[f'g{r_g}{c_g}'] = val_g

            # AO values
            val_ao1 = str(field_map.get(f'g{r_g}_5_ao1') or field_map.get(f'g{r_g}5_ao1') or field_map.get(f'g{r_g}_5') or field_map.get(f'g{r_g}5') or '')
            val_ao2 = str(field_map.get(f'g{r_g}_5_ao2') or field_map.get(f'g{r_g}5_ao2') or field_map.get(f'g{r_g}_6') or field_map.get(f'g{r_g}6') or '')
            val_ao3 = str(field_map.get(f'g{r_g}_5_ao3') or field_map.get(f'g{r_g}5_ao3') or field_map.get(f'g{r_g}_7') or field_map.get(f'g{r_g}7') or '')

            field_map[f'g{r_g}5'] = val_ao1
            field_map[f'g{r_g}_5'] = val_ao1
            field_map[f'g{r_g}5_ao1'] = val_ao1
            field_map[f'g{r_g}_5_ao1'] = val_ao1

            field_map[f'g{r_g}5_ao2'] = val_ao2
            field_map[f'g{r_g}_5_ao2'] = val_ao2

            field_map[f'g{r_g}5_ao3'] = val_ao3
            field_map[f'g{r_g}_5_ao3'] = val_ao3

        field_map['inst'] = str(field_map.get('instrument_used') or data.get('instrument_used') or '')
        field_map['tb'] = str(field_map.get('tested_by') or data.get('tested_by') or '')
        field_map['ab'] = str(field_map.get('approved_by') or data.get('approved_by') or '')
        field_map['dt'] = str(field_map.get('date') or data.get('date') or '')

    # Format 2-column header paragraphs so Customer Name is EXACTLY vertically aligned above Conveyor No / Tag No / Material
    if report_type == 'belt_scale':
        job_val = str(field_map.get('job') or field_map.get('job_no') or data.get('job_no') or '').strip()
        cust_val = str(field_map.get('cust') or field_map.get('customer') or field_map.get('customer_name') or data.get('customer') or '').strip()
        cap_val = str(field_map.get('cap') or field_map.get('capacity') or data.get('capacity') or '').strip()
        if cap_val and not any(u in cap_val.upper() for u in ['TPH', 'KG', 'TON']):
            cap_val = f"{cap_val} TPH"
        cn_val = str(field_map.get('cn') or field_map.get('conveyor_no') or data.get('conveyor_no') or '').strip()

        cust_str = str(cust_val or "")
        if len(cust_str) > 55:
            hdr_tab, hdr_pt = 3600, 9.5
        elif len(cust_str) > 40:
            hdr_tab, hdr_pt = 3900, 10.0
        else:
            hdr_tab, hdr_pt = 4200, 11.0

        for p in doc.paragraphs:
            txt = p.text
            if "JOB No:" in txt or "JOB NO:" in txt or "{{job}}" in txt:
                set_para_with_tab(p, f"JOB No: {job_val}", f"CUSTOMER NAME: {cust_val}", hdr_tab, hdr_pt, bold=False)
            elif "CAPACITY" in txt or "{{cap}}" in txt:
                set_para_with_tab(p, f"CAPACITY: {cap_val}", f"CONVEYOR NO: {cn_val}", hdr_tab, hdr_pt, bold=False)

    elif report_type in ['weigh_feeder', 'loss_in_weigh_feeder', 'batching_system']:
        job_val = str(field_map.get('job') or field_map.get('job_no') or data.get('job_no') or '').strip()
        cust_val = str(field_map.get('cust') or field_map.get('customer') or field_map.get('customer_name') or data.get('customer') or '').strip()
        cap_val = str(field_map.get('cap') or field_map.get('capacity') or data.get('capacity') or '').strip()
        mat_val = str(field_map.get('material') or field_map.get('material_name') or data.get('material') or '').strip()

        cust_str = str(cust_val or "")
        if len(cust_str) > 55:
            hdr_tab, hdr_pt = 3600, 9.5
        elif len(cust_str) > 40:
            hdr_tab, hdr_pt = 3900, 10.0
        else:
            hdr_tab, hdr_pt = 4200, 11.0

        for p in doc.paragraphs:
            txt = p.text
            if "JOB No" in txt and "CUSTOMER" in txt:
                set_para_with_tab(p, f"JOB No: {job_val}", f"CUSTOMER: {cust_val}", hdr_tab, hdr_pt, bold=False)
            elif "CAPACITY" in txt and "MATERIAL" in txt:
                unit = "TPH" if report_type != 'batching_system' else "Kg/Tonnes"
                if cap_val and not any(u in cap_val.upper() for u in ['TPH', 'KG', 'TON']):
                    cap_val = f"{cap_val} {unit}"
                set_para_with_tab(p, f"CAPACITY: {cap_val}", f"MATERIAL: {mat_val}", hdr_tab, hdr_pt, bold=False)

    # Populate templates that do not use {{placeholder}} markers (all 16 specialized forms)
    _apply_template_specific_fillers(doc, report_type, data, field_map)

    # 1. Paragraph replacement
    for p in doc.paragraphs:
        replace_in_paragraph(p, field_map)

    # 2. Table cell replacement
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, field_map)

    # 4. Clean up any unhandled template tags
    for p in doc.paragraphs:
        clean_unhandled_tags(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    clean_unhandled_tags(p)

    # 5. Belt Scale layout guards: maintain 1-page geometry without altering document margins
    if report_type == 'belt_scale':
        if len(doc.tables) > 1 and "Check for short" in doc.tables[1].cell(0, 0).text:
            sc_val = str(field_map.get("sc") or data.get("short_check") or "CHECKED AND FOUND OK").strip()
            doc.tables[1].cell(0, 0).text = f"Check for short between Line & Neutral, Neutral & Earth, Line & Earth: {sc_val}"
            for cp in doc.tables[1].cell(0, 0).paragraphs:
                for r in cp.runs:
                    r.font.size = Pt(11)

        # Rebalance Table 0 columns for 8 LC so both serial columns fit on 4 lines without wrapping
        if lc_system == '8' and len(doc.tables) > 0:
            t0 = doc.tables[0]
            col_w = [2592, 3744, 2304, 2736]
            for r in t0.rows:
                if len(r.cells) >= 4:
                    for c_i, w in enumerate(col_w):
                        r.cells[c_i].width = int(w * 635)
                        r.cells[c_i]._tc.tcPr.tcW.w = w
            if hasattr(t0._tbl, 'tblGrid') and t0._tbl.tblGrid is not None and len(t0._tbl.tblGrid.gridCol_lst) >= 4:
                for c_i, w in enumerate(col_w):
                    t0._tbl.tblGrid.gridCol_lst[c_i].set(qn('w:w'), str(w))

            # Ensure left tab stop at 1750 dxa and 11pt font on the 4 serial lines
            if len(t0.rows) > 1 and len(t0.rows[1].cells) > 1:
                c1 = t0.rows[1].cells[1]
                for p_idx in range(4, min(8, len(c1.paragraphs))):
                    p_ser = c1.paragraphs[p_idx]
                    pPr = p_ser._p.get_or_add_pPr()
                    for old_tabs in pPr.findall(qn('w:tabs')):
                        pPr.remove(old_tabs)
                    tabs_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="left" w:pos="1750"/></w:tabs>')
                    pPr.append(tabs_xml)
                    for r in p_ser.runs:
                        r.font.size = Pt(11)

        # Ensure 3 AO calibration grid (Table 3) has enough width in AO3 column to prevent 5-char wrap
        if num_ao == '3' and len(doc.tables) > 3:
            t3 = doc.tables[3]
            for r in t3.rows[1:]:
                if len(r.cells) >= 8:
                    r.cells[4].width = int(1739 * 635)
                    r.cells[4]._tc.tcPr.tcW.w = 1739
                    r.cells[7].width = int(809 * 635)
                    r.cells[7]._tc.tcPr.tcW.w = 809

        # Enforce standard 11pt font size on all runs that defaulted to Normal style (12pt)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            if r.font.size is None:
                                r.font.size = Pt(11)
        for p in doc.paragraphs:
            for r in p.runs:
                if r.font.size is None:
                    r.font.size = Pt(11)

    # Output path is passed from app.py
    doc.save(output_path)

    # Mirror DOCX directly into \\192.168.100.248\prdndata\TEST REPORT
    if os.path.exists(output_path):
        try:
            if os.path.exists(SERVER_TEST_REPORT_DIR):
                import shutil
                shutil.copyfile(output_path, os.path.join(SERVER_TEST_REPORT_DIR, os.path.basename(output_path)))
        except Exception:
            pass

    pdf_path = output_path.replace(".docx", ".pdf")
    convert_docx_to_pdf(output_path, pdf_path)

    return pdf_path
