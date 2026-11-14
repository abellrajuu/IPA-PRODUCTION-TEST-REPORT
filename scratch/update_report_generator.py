file_path = "report_generator_python.py"
with open(file_path, "r", encoding="utf-8") as f:
    text = f.read()

target = """    span_val = 10.0
    unit_str = "Kg" if is_load_calib else "mV"
    raw_cap_str = get_val("capacity", "cap", default="")"""

replacement = """    span_val = 10.0
    unit_str = "Kg" if is_load_calib else "mV"
    raw_cap_str = get_val("spec_4", "act_4", "capacity", "cap", default="")"""

norm_text = text.replace("\r\n", "\n")
norm_target = target.replace("\r\n", "\n")
norm_replacement = replacement.replace("\r\n", "\n")

if norm_target in norm_text:
    norm_text = norm_text.replace(norm_target, norm_replacement, 1)
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(norm_text)
    print("Updated report_generator_python.py successfully!")
else:
    print("Target not found in report_generator_python.py")
