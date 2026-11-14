import docx, os

docs_to_check = [
    "DOC/Belt Scale.docx",
    "DOC/PP-05_06 Digital Indicator.docx",
    "DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM.docx",
    "DOC/Signal Conditioner.docx",
    "DOC/Miscellaneous report.docx"
]

for dpath in docs_to_check:
    if not os.path.exists(dpath):
        print(f"File not found: {dpath}")
        continue
    doc = docx.Document(dpath)
    print(f"\n==========================================")
    print(f"DOCUMENT: {dpath}")
    print(f"==========================================")
    print("--- PARAGRAPHS (First 15) ---")
    for i, p in enumerate(doc.paragraphs[:15]):
        # check tabs in paragraph format
        tabs_info = [(t.position.pt, str(t.alignment)) for t in p.paragraph_format.tab_stops]
        print(f"P{i}: text='{p.text}' | tabs={tabs_info}")
    
    print(f"--- TABLES: {len(doc.tables)} tables ---")
    for t_i, t in enumerate(doc.tables):
        print(f"Table {t_i}: rows={len(t.rows)}, cols={len(t.columns)}")
        # Print first 2 rows
        for r_i, r in enumerate(t.rows[:2]):
            cells_text = [c.text.replace('\n', ' ')[:30] for c in r.cells]
            print(f"  R{r_i}: {cells_text}")
