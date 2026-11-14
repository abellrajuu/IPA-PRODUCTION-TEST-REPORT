import sys, os
sys.path.insert(0, os.path.abspath('.'))
from report_generator_python import generate_digital_indicator_record
import docx

test_data = {
    "job_no": "TEST-67KG",
    "customer": "Test Customer Ltd",
    "calib_mode": "load",
    "lc_system": "4",
    "capacity": "150T",
    "spec_4": "67.0KG",
    "act_4": "67.0KG",
}

out_docx = "scratch/test_67kg_output.docx"
generate_digital_indicator_record(test_data, out_docx)

doc = docx.Document(out_docx)
print("Doc tables count:", len(doc.tables))
t2 = doc.tables[2]
print("Table 2 Row 4 (Spec, Act):", [t2.cell(4, 2).text, t2.cell(4, 3).text])

t3 = doc.tables[3]
print("Table 3 (Matrix) Rows 1 to 5:")
for r in range(1, 6):
    row_vals = [c.text for c in t3.rows[r].cells[:6]]
    print(f"  Step {r}: {row_vals}")
