import docx

doc = docx.Document("DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM.docx")
print("=== PARAGRAPHS ===")
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"P{i}: {repr(p.text)}")

print("\n=== TABLES ===")
for t_i, t in enumerate(doc.tables):
    print(f"\nTable {t_i}: {len(t.rows)} rows x {len(t.columns)} cols")
    for r_i, r in enumerate(t.rows):
        row_txt = [c.text.replace('\n', ' ') for c in r.cells]
        print(f"  R{r_i}: {row_txt}")
