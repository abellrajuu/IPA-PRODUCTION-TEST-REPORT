import os, re
import docx

p_html = r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\inprocess.html"
with open(p_html, "r", encoding="utf-8", errors="ignore") as f:
    text = f.read()

print("=== inprocess.html ===")
inputs = re.findall(r'<input[^>]*name=["\']([^"\']+)["\']', text)
print("Input names:", inputs)

p_docx = r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\backend_docx_templates\PP-05_16 Inprocess.docx"
if os.path.exists(p_docx):
    doc = docx.Document(p_docx)
    print("\n=== PP-05_16 Inprocess.docx ===")
    for p in doc.paragraphs:
        if p.text.strip():
            print("P:", p.text.strip())
    for i, t in enumerate(doc.tables):
        print(f"\n--- Table {i} ({len(t.rows)} rows x {len(t.columns)} cols) ---")
        for r_idx, r in enumerate(t.rows[:8]):
            row_txt = [c.text.strip().replace('\n', ' ') for c in r.cells]
            print(f"R{r_idx}:", " | ".join(row_txt[:8]))
