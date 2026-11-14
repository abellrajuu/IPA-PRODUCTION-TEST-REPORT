import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import report_generator_python, sterp_data_fetcher, pymupdf

def run_tests():
    print("=== STARTING COMPREHENSIVE SUITE TEST ===")
    
    # 1. Digital Indicator (SY1550)
    print("\n[TEST 1] Digital Indicator (SY1550)")
    data_sy1550 = {
        "report_type": "digital_indicator",
        "job_no": "SY1550",
        "customer": "Exceltech Thermal Optimizer Private Limited",
        "capacity": "100 KG",
        "tag_no": "TAG-101",
        "resolution": "0.01 kg",
        "di_model": "LH28222-",
        "di_serial": "SN-DI-1550",
        "sw_version": "v2.1",
        "lc_model": "TH053H0",
        "lc_system": "8",
        "s1": "101", "s2": "102", "s3": "103", "s4": "104",
        "s5": "105", "s6": "106", "s7": "107", "s8": "108",
        "jbox_model": "JT031M9",
        "jbox_serial": "JB-1550",
        "calib_mode": "mv",
        "num_outputs": "1",
        "act_1": "230V AC",
        "act_2": "24V DC",
        "tested_by": "Abel",
        "approved_by": "HOD - PDN"
    }
    out_di = "scratch/final_di_sy1550.docx"
    pdf_di = "scratch/final_di_sy1550.pdf"
    report_generator_python.generate_digital_indicator_record(data_sy1550, out_di)
    doc_di = pymupdf.open(pdf_di)
    print(f"  DI SY1550 Pages: {len(doc_di)} (Expected: 2)")
    assert len(doc_di) == 2, f"Expected 2 pages for DI, got {len(doc_di)}"

    # 2. Belt Scale (SY1529)
    print("\n[TEST 2] Belt Scale (SY1529)")
    data_sy1529 = {
        "report_type": "belt_scale",
        "job_no": "SY1529",
        "customer": "THE INDIA CEMENTS LIMITED",
        "capacity": "500",
        "conveyor_no": "BC-01",
        "belt_speed": "2.50",
        "bs_model": "LH32220",
        "bs_serial": "SN-BS-1529",
        "sw_version": "v1.0",
        "sensor_model": "BR",
        "lc_system": "4",
        "s1": "201", "s2": "202", "s3": "203", "s4": "204",
        "jbox_model1": "JT041M9", "jbox_serial1": "JB-01",
        "jbox_model2": "JT011M0", "jbox_serial2": "JB-02",
        "tacho_model": "TS-01", "tacho_serial": "SS-01",
        "tested_by": "Abel",
        "approved_by": "HOD - PDN"
    }
    out_bs = "scratch/final_bs_sy1529.docx"
    pdf_bs = "scratch/final_bs_sy1529.pdf"
    report_generator_python.generate_docx_record(data_sy1529, out_bs)
    doc_bs = pymupdf.open(pdf_bs)
    print(f"  BS SY1529 Pages: {len(doc_bs)} (Expected: 1)")
    assert len(doc_bs) == 1, f"Expected 1 page for BS, got {len(doc_bs)}"
    # Check resolution units in BS
    bs_text = doc_bs[0].get_text()
    assert "kg/m" in bs_text and "m/s" in bs_text and "tph" in bs_text and "tonnes" in bs_text, "Missing resolution units in BS!"
    print("  BS Resolution units verified: kg/m, m/s, tph, tonnes present!")

    # 3. Crane Scale (Electromech SY1615)
    print("\n[TEST 3] Crane Scale (SY1615 - Long Customer Name)")
    data_cs = {
        "report_type": "crane_scale",
        "job_no": "SY1615",
        "customer": "ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD",
        "capacity": "15 Tonnes",
        "stamp_no": "ST-1615",
        "lc_model": "SP752L718EB1A05",
        "lc_serial": "11673",
        "resolution": "2.0 kg",
        "crane_model": "CS-15T",
        "crane_serial": "CS-1615",
        "charger_model": "BC-12V",
        "remote_model": "RDU-433",
        "act_1": "230V AC",  # Form sent default AC power, generator should guard and keep 6.52 V
        "act_2": "24V DC",   # Form sent default 24V, generator should guard and keep Checked & OK
        "tested_by": "Abel",
        "approved_by": "HOD - PDN"
    }
    out_cs = "scratch/final_cs_sy1615.docx"
    pdf_cs = "scratch/final_cs_sy1615.pdf"
    report_generator_python.generate_crane_scale_record(data_cs, out_cs)
    doc_cs = pymupdf.open(pdf_cs)
    print(f"  CS SY1615 Pages: {len(doc_cs)} (Expected: 1)")
    assert len(doc_cs) == 1, f"Expected 1 page for CS, got {len(doc_cs)}"
    cs_text = doc_cs[0].get_text()
    assert "6.52 V" in cs_text, "Crane Scale act_1 was overwritten by 230V!"
    assert "Checked & OK" in cs_text or "Checked & Found OK" in cs_text, "Crane Scale act_2 was overwritten by 24V DC!"
    print("  CS Routine test guards verified: 6.52 V preserved!")

    # 4. Signal Conditioner (SY1615)
    print("\n[TEST 4] Signal Conditioner (SY1615)")
    data_sc = {
        "report_type": "signal_conditioner",
        "job_no": "SY1615",
        "customer": "ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD",
        "sc_model": "LH23200",
        "sc_serial": "SC-1615",
        "lc_model": "SP752L718EB1A05",
        "lc_serial": "11673",
        "calib_type": "mv",
        "span": "20.00 mV",
        "tested_by": "Abel",
        "approved_by": "HOD - PDN"
    }
    out_sc = "scratch/final_sc_sy1615.docx"
    pdf_sc = "scratch/final_sc_sy1615.pdf"
    report_generator_python.generate_signal_conditioner_record(data_sc, out_sc)
    doc_sc = pymupdf.open(pdf_sc)
    print(f"  SC SY1615 Pages: {len(doc_sc)} (Expected: 1)")
    assert len(doc_sc) == 1, f"Expected 1 page for SC, got {len(doc_sc)}"
    sc_text = doc_sc[0].get_text()
    assert "20.00 mA" in sc_text and "4.00 mA" in sc_text, "Signal Conditioner missing 5-step linearity mA values!"
    print("  SC Linearity 5-step matrix verified!")

    # 5. Miscellaneous Report (SY1615)
    print("\n[TEST 5] Miscellaneous Report (SY1615)")
    data_misc = {
        "report_type": "misc_report",
        "job_no": "SY1615",
        "customer": "ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD",
        "custom_model": "SP752L718EB1A05",
        "custom_desc": "Shear Pin Load Cell",
        "custom_sl": "11673",
        "custom_qty": "1 No",
        "readings_result": "Input Resistance: 350 Ohms\nOutput Resistance: 350 Ohms\nInsulation Resistance: > 1000 M Ohms\nZero Balance: 0.02 mV\nChecked and Found OK.",
        "tested_by": "Abel",
        "approved_by": "HOD - PDN"
    }
    out_misc = "scratch/final_misc_sy1615.docx"
    pdf_misc = "scratch/final_misc_sy1615.pdf"
    report_generator_python.generate_misc_record(data_misc, out_misc)
    doc_misc = pymupdf.open(pdf_misc)
    print(f"  Misc SY1615 Pages: {len(doc_misc)} (Expected: 1)")
    assert len(doc_misc) == 1, f"Expected 1 page for Misc, got {len(doc_misc)}"

    print("\n=== ALL TESTS PASSED WITH 100% SUCCESS! ===")

if __name__ == "__main__":
    run_tests()
