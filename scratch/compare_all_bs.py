import docx

pairs = [
    ("DOC/Belt Scale.docx", "backend_docx_templates/PP-05_03 Belt Scale.docx"),
    ("DOC/Belt Scale with 1 current output with 8 load cell system.docx", "backend_docx_templates/PP-05_03 Belt Scale - 8 LC.docx"),
    ("DOC/Belt Scale with 2 current ouput.docx", "backend_docx_templates/PP-05_03 Belt Scale - 2 AO.docx"),
    ("DOC/Belt Scale with 2 current output with 8 load cell system.docx", "backend_docx_templates/PP-05_03 Belt Scale - 2 AO - 8 LC.docx"),
    ("DOC/Belt Scale with 3 current ouput.docx", "backend_docx_templates/PP-05_03 Belt Scale - 3 AO.docx"),
    ("DOC/Belt Scale with 3 current output with 8 load cell system.docx", "backend_docx_templates/PP-05_03 Belt Scale - 3 AO - 8 LC.docx"),
]

for og, be in pairs:
    d1 = docx.Document(og)
    d2 = docx.Document(be)
    print(f"\nComparing {og} vs {be}")
    print(f"  OG paras={len(d1.paragraphs)}, tables={len(d1.tables)} | BE paras={len(d2.paragraphs)}, tables={len(d2.tables)}")
    # Compare table row/col counts
    for t_i in range(len(d1.tables)):
        t1 = d1.tables[t_i]
        t2 = d2.tables[t_i]
        if (len(t1.rows), len(t1.columns)) != (len(t2.rows), len(t2.columns)):
            print(f"    Table {t_i} MISMATCH: OG={len(t1.rows)}x{len(t1.columns)} vs BE={len(t2.rows)}x{len(t2.columns)}")
        else:
            print(f"    Table {t_i} MATCHES: {len(t1.rows)}x{len(t1.columns)}")
