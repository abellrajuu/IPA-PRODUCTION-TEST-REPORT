import docx

d1 = docx.Document("DOC/Belt Scale.docx")
d2 = docx.Document("backend_docx_templates/PP-05_03 Belt Scale.docx")

print("--- BELT SCALE COMPARISON ---")
print("DOC paragraphs:", len(d1.paragraphs), "backend paragraphs:", len(d2.paragraphs))
print("DOC tables:", len(d1.tables), "backend tables:", len(d2.tables))

for i in range(len(d1.paragraphs)):
    t1 = d1.paragraphs[i].text.strip()
    t2 = d2.paragraphs[i].text.strip()
    if t1 != t2:
        print(f"P{i}:\n  DOC:     '{t1}'\n  BACKEND: '{t2}'")

for t_i in range(len(d1.tables)):
    t1 = d1.tables[t_i]
    t2 = d2.tables[t_i]
    print(f"Table {t_i}: DOC rows={len(t1.rows)}, cols={len(t1.columns)} | BACKEND rows={len(t2.rows)}, cols={len(t2.columns)}")
