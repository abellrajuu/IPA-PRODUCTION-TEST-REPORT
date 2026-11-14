import os, docx

p1 = r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\backend_docx_templates\PP-05_02 Job Traveller Card.docx"
p2 = r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\backend_docx_templates\PP-05_16 Inprocess.docx"

for path in [p1, p2]:
    print(f"\n=================== {os.path.basename(path)} ===================")
    doc = docx.Document(path)
    for p in doc.paragraphs:
        if p.text.strip():
            print("P:", p.text.strip())
    for i, t in enumerate(doc.tables):
        print(f"\n--- Table {i} ({len(t.rows)}x{len(t.columns)}) ---")
        for r_idx, r in enumerate(t.rows):
            cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
            print(f"R{r_idx}:", " | ".join(cells[:8]))
