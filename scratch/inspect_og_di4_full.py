import docx

doc = docx.Document("DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM.docx")
print("Total Paragraphs:", len(doc.paragraphs))
for i, p in enumerate(doc.paragraphs):
    if p.text.strip():
        print(f"P{i:2}: {repr(p.text)}")

print("\nTotal Tables:", len(doc.tables))
for t_idx, t in enumerate(doc.tables):
    print(f"\n--- TABLE {t_idx} ({len(t.rows)} rows x {len(t.columns)} cols) ---")
    for r_idx, r in enumerate(t.rows):
        txts = [c.text.replace('\n', ' // ') for c in r.cells]
        print(f"  R{r_idx}: {txts}")
