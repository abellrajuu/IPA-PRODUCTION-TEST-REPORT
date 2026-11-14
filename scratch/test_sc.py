import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx
from docx.shared import Pt
import report_generator_python, pymupdf

test_data = {
    "report_type": "signal_conditioner",
    "job_no": "SY1615",
    "customer": "ELECTROMECH MATERIAL HANDLING SYSTEMS (INDIA)PVT.LTD",
    "sc_model": "SC-100",
    "sc_serial": "SC-2026-001",
    "lc_model": "SP752L",
    "lc_serial": "11673",
    "calib_type": "mv",
    "span": "20.00 mV",
    "tare": "0.00 mV",
    "exc_v": "10.0 VDC",
    "annunciation": "OK",
    "pct_error": "< 0.05%",
    "multimeter_used": "FLUKE 179",
    "tested_by": "Abel",
    "approved_by": "HOD - PDN",
}

out_docx = "scratch/test_sc_electromech.docx"
out_pdf = "scratch/test_sc_electromech.pdf"
report_generator_python.generate_signal_conditioner_record(test_data, out_docx)
doc_pdf = pymupdf.open(out_pdf)
print("SC PDF Pages:", len(doc_pdf))
for b in doc_pdf[0].get_text("dict")["blocks"]:
    if b.get("type") == 0:
        for l in b["lines"]:
            y = l["bbox"][1]
            if 60 < y < 200:
                print(f"Line y={y:.1f}:")
                for s in l["spans"]:
                    print(f"  x={s['bbox'][0]:.1f} font={s['font']} size={s['size']:.1f} text={repr(s['text'])}")
