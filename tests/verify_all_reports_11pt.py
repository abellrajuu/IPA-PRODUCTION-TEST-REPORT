import os
import sys
import docx
import pymupdf as fitz

BASE_DIR = r"C:\Users\Admin\Desktop\PDN TEST REPORT"
sys.path.insert(0, BASE_DIR)
import report_generator_python as rg

OUT_DIR = os.path.join(BASE_DIR, "tests", "test_all_11pt_results")
os.makedirs(OUT_DIR, exist_ok=True)

test_cases = [
    # 1. Belt Scale 8LC variants
    {
        "name": "BS_8LC_1AO",
        "data": {
            "report_type": "belt_scale",
            "job_no": "JOB-BS-8LC-1AO",
            "customer": "Tata Steel Limited - Raw Material Division",
            "capacity": "1200 TPH",
            "conveyor_no": "BC-201A",
            "belt_speed": "2.50 m/s",
            "bs_model": "BS-600",
            "bs_serial": "BS-2026-99",
            "sw_version": "v4.1.0",
            "lc_model": "LC-CZ-50T",
            "lc_system": "8",
            "num_outputs": "1",
            "s1": "LC-2026-001", "s2": "LC-2026-002", "s3": "LC-2026-003", "s4": "LC-2026-004",
            "s5": "LC-2026-005", "s6": "LC-2026-006", "s7": "LC-2026-007", "s8": "LC-2026-008",
            "ss_model": "SS-100", "ss_serial": "SS-881",
            "as_model": "AS-50", "as_serial": "AS-11",
            "jb_1_model": "JB-8P", "jb_1_serial": "JB-001",
            "grid_data": [[f"{i*100}", f"{i*20}", "2.50", f"{i*100}", f"{i*50}", f"{4+i*1.6:.2f}"] for i in range(11)]
        },
        "expected_pages": 1
    },
    {
        "name": "BS_8LC_2AO",
        "data": {
            "report_type": "belt_scale",
            "job_no": "JOB-BS-8LC-2AO",
            "customer": "Tata Steel Limited",
            "capacity": "1200 TPH",
            "conveyor_no": "BC-201A",
            "belt_speed": "2.50 m/s",
            "bs_model": "BS-600",
            "bs_serial": "BS-2026-99",
            "sw_version": "v4.1.0",
            "lc_model": "LC-CZ-50T",
            "lc_system": "8",
            "num_outputs": "2",
            "s1": "LC-2026-001", "s2": "LC-2026-002", "s3": "LC-2026-003", "s4": "LC-2026-004",
            "s5": "LC-2026-005", "s6": "LC-2026-006", "s7": "LC-2026-007", "s8": "LC-2026-008",
            "ss_model": "SS-100", "ss_serial": "SS-881",
            "as_model": "AS-50", "as_serial": "AS-11",
            "jb_1_model": "JB-8P", "jb_1_serial": "JB-001",
            "grid_data": [[f"{i*100}", f"{i*20}", "2.50", f"{i*100}", f"{i*50}", f"{4+i*1.6:.2f}", f"{4+i*1.6:.2f}"] for i in range(11)]
        },
        "expected_pages": 1
    },
    {
        "name": "BS_8LC_3AO",
        "data": {
            "report_type": "belt_scale",
            "job_no": "JOB-BS-8LC-3AO",
            "customer": "Tata Steel Limited",
            "capacity": "1200 TPH",
            "conveyor_no": "BC-201A",
            "belt_speed": "2.50 m/s",
            "bs_model": "BS-600",
            "bs_serial": "BS-2026-99",
            "sw_version": "v4.1.0",
            "lc_model": "LC-CZ-50T",
            "lc_system": "8",
            "num_outputs": "3",
            "s1": "LC-2026-001", "s2": "LC-2026-002", "s3": "LC-2026-003", "s4": "LC-2026-004",
            "s5": "LC-2026-005", "s6": "LC-2026-006", "s7": "LC-2026-007", "s8": "LC-2026-008",
            "ss_model": "SS-100", "ss_serial": "SS-881",
            "as_model": "AS-50", "as_serial": "AS-11",
            "jb_1_model": "JB-8P", "jb_1_serial": "JB-001",
            "grid_data": [[f"{i*100}", f"{i*20}", "2.50", f"{i*100}", f"{i*50}", f"{4+i*1.6:.2f}", f"{4+i*1.6:.2f}", f"{4+i*1.6:.2f}"] for i in range(11)]
        },
        "expected_pages": 1
    },
    # 2. Belt Scale 4LC
    {
        "name": "BS_4LC_1AO",
        "data": {
            "report_type": "belt_scale",
            "job_no": "JOB-BS-4LC-1AO",
            "customer": "UltraTech Cement",
            "capacity": "500 TPH",
            "conveyor_no": "CV-01",
            "belt_speed": "1.80 m/s",
            "bs_model": "BS-400",
            "bs_serial": "BS-2026-12",
            "sw_version": "v4.1.0",
            "lc_model": "LC-CZ-20T",
            "lc_system": "4",
            "num_outputs": "1",
            "s1": "LC-01", "s2": "LC-02", "s3": "LC-03", "s4": "LC-04",
            "ss_model": "SS-100", "ss_serial": "SS-12",
            "jb_1_model": "JB-4P", "jb_1_serial": "JB-12",
            "grid_data": [[f"{i*50}", f"{i*10}", "1.80", f"{i*50}", f"{i*25}", f"{4+i*1.6:.2f}"] for i in range(11)]
        },
        "expected_pages": 1
    },
    # 3. Digital Indicator 8LC Load
    {
        "name": "DI_8LC_Load",
        "data": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-8LC-LOAD",
            "customer": "Tata Steel Limited",
            "capacity": "100 Tonnes",
            "tag_no": "DI-TAG-01",
            "di_model": "DI-3000",
            "di_serial": "SN-DI-8821",
            "sw_version": "v3.2.0",
            "lc_model": "LC-CZ50T",
            "lc_system": "8",
            "calib_mode": "load",
            "current_output": "yes",
            "s1": "LC-SN-001", "s2": "LC-SN-002", "s3": "LC-SN-003", "s4": "LC-SN-004",
            "s5": "LC-SN-005", "s6": "LC-SN-006", "s7": "LC-SN-007", "s8": "LC-SN-008",
            "remote_model": "N/A", "remote_serial": "N/A",
            "jbox_model": "JB-8P", "jbox_serial": "JB-SN-441",
            "grid_data": [[f"{i*10}%"] + [f"{i*100}" for _ in range(8)] + [f"{4+i*1.6:.2f}"] for i in range(12)]
        },
        "expected_pages": 2
    },
    # 4. Digital Indicator 8LC mV
    {
        "name": "DI_8LC_mV",
        "data": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-8LC-MV",
            "customer": "JSW Steel Ltd",
            "capacity": "80 Tonnes",
            "tag_no": "DI-TAG-02",
            "di_model": "DI-3000",
            "di_serial": "SN-DI-9901",
            "sw_version": "v3.2.0",
            "lc_model": "LC-CZ50T",
            "lc_system": "8",
            "calib_mode": "mv",
            "current_output": "yes",
            "s1": "LC-SN-101", "s2": "LC-SN-102", "s3": "LC-SN-103", "s4": "LC-SN-104",
            "s5": "LC-SN-105", "s6": "LC-SN-106", "s7": "LC-SN-107", "s8": "LC-SN-108",
            "jbox_model": "JB-8P", "jbox_serial": "JB-SN-442",
            "grid_data": [[f"{i*10}%"] + [f"{i*2.0:.2f}" for _ in range(8)] + [f"{4+i*1.6:.2f}"] for i in range(12)]
        },
        "expected_pages": 2
    },
    # 5. Signal Conditioner Load Cell
    {
        "name": "SC_LoadCell",
        "data": {
            "report_type": "signal_conditioner",
            "job_no": "JOB-SC-LOAD",
            "customer": "Automation Systems Ltd",
            "sc_model": "SC-4020",
            "sc_serial": "SN-SC-8899",
            "lc_model": "LC-CZ50T",
            "lc_serial": "LC-SN-001",
            "calib_type": "load",
            "span": "10.00",
            "tare": "0.00",
            "exc_v": "10.0V",
            "annunciation": "OK",
            "pct_error": "0.01%",
            "multimeter_used": "Fluke 87V",
            "tested_by": "Inspector",
            "approved_by": "HOD-PDN",
            "grid_data": [[f"{i*20}%", f"{4+i*1.6:.2f}"] for i in range(6)]
        },
        "expected_pages": 1
    },
    # 6. Crane Scale
    {
        "name": "Crane_Scale",
        "data": {
            "report_type": "crane_scale",
            "job_no": "JOB-CS-2026",
            "customer": "Heavy Forge Industries",
            "capacity": "10 Tonnes",
            "stamp_no": "STAMP-2026-01",
            "lc_model": "LC-CWS-10T",
            "lc_serial": "SN-LC-001",
            "resolution": "1.0 kg",
            "crane_model": "CS-10T",
            "crane_serial": "SN-CS-001",
            "charger_model": "BC-12V",
            "charger_serial": "CH-001",
            "remote_model": "RDU-RF-433",
            "remote_serial": "RD-001",
            "hook_serial": "HK-001",
            "shackle_serial": "SH-001",
            "act_1": "6.52 V (PASS)",
            "act_2": "Checked & OK",
            "act_3": "Working Normal",
            "tested_by": "Inspector",
            "approved_by": "HOD-PDN"
        },
        "expected_pages": 1
    },
    # 7. Misc Reports
    {
        "name": "Misc_LoadCell",
        "data": {
            "report_type": "misc_report",
            "misc_type": "load_cell",
            "chk_loadcell": "true",
            "job_no": "JOB-MISC-LC",
            "customer": "Tata Steel Limited",
            "lc_serial_1": "LC-01", "lc_in_res_1": "350.2", "lc_out_res_1": "350.0", "lc_unbalance_1": "0.01",
            "lc_serial_2": "LC-02", "lc_in_res_2": "350.1", "lc_out_res_2": "350.2", "lc_unbalance_2": "0.02",
            "tested_by": "Tester",
            "approved_by": "HOD-PDN"
        },
        "expected_pages": 1
    },
    {
        "name": "Misc_RemoteDisplay",
        "data": {
            "report_type": "misc_report",
            "misc_type": "remote_display",
            "chk_remote": "true",
            "job_no": "JOB-MISC-RD",
            "customer": "Tata Steel Limited",
            "remote_model": "RD-100",
            "remote_serial": "RD-001",
            "tested_by": "Tester",
            "approved_by": "HOD-PDN"
        },
        "expected_pages": 1
    },
    {
        "name": "Misc_JBox_Multiple",
        "data": {
            "report_type": "misc_report",
            "misc_type": "jb_multiple",
            "chk_jbox": "true",
            "job_no": "JOB-MISC-JB-MULTI",
            "customer": "Tata Steel Limited",
            "variant_model_1": "JB-04",
            "variant_serial_1": "JB-001",
            "variant_model_2": "JB-08",
            "variant_serial_2": "JB-002",
            "tested_by": "Tester",
            "approved_by": "HOD-PDN"
        },
        "expected_pages": 1
    }
]

print("=== STARTING REPORT VERIFICATION AT 11PT ===")
all_pass = True

for tc in test_cases:
    name = tc["name"]
    data = tc["data"]
    exp = tc["expected_pages"]
    docx_path = os.path.join(OUT_DIR, f"{name}.docx")
    pdf_path = os.path.join(OUT_DIR, f"{name}.pdf")
    
    try:
        rg.generate_docx_record(data, docx_path)
        doc = fitz.open(pdf_path)
        actual_pages = len(doc)
        
        status = "PASS" if actual_pages == exp else "FAIL"
        if status == "FAIL":
            all_pass = False
        print(f"[{status}] {name:20}: Pages={actual_pages} (Expected={exp})")
        
        # Save snapshot of page 0
        doc[0].get_pixmap(dpi=150).save(os.path.join(OUT_DIR, f"{name}_p0.png"))
        if actual_pages > 1:
            doc[1].get_pixmap(dpi=150).save(os.path.join(OUT_DIR, f"{name}_p1.png"))
            
    except Exception as e:
        print(f"[ERROR] {name:20}: {e}")
        all_pass = False

print(f"\nOVERALL RESULT: {'ALL TESTS PASSED' if all_pass else 'SOME TESTS FAILED'}")
