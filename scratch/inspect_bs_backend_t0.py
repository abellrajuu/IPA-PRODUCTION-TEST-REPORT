import docx

d = docx.Document("backend_docx_templates/PP-05_03 Belt Scale.docx")
t0 = d.tables[0]
print(f"backend Belt Scale Table 0: rows={len(t0.rows)}")
for r_i, r in enumerate(t0.rows):
    for c_i, c in enumerate(r.cells):
        print(f"Cell ({r_i}, {c_i}): {repr(c.text.replace(chr(10), ' // '))}")
