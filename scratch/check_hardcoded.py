import os, re

base = r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates"
found = False
for f in os.listdir(base):
    if f.endswith(".html"):
        path = os.path.join(base, f)
        with open(path, "r", encoding="utf-8") as fp:
            for l in fp:
                if 'name="job_no"' in l or 'id="job_no"' in l:
                    if 'value=' in l and 'value=""' not in l:
                        print(f"Hardcoded in {f}: {l.strip()}")
                        found = True
if not found:
    print("Clean: No HTML template has any hardcoded job numbers!")
