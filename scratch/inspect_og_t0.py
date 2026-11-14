import docx

for doc_path in ["DOC/Belt Scale.docx", "DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM.docx", "DOC/Signal Conditioner.docx"]:
    d = docx.Document(doc_path)
    print(f"\n==========================================")
    print(f"DOCUMENT: {doc_path} - TABLE 0")
    print(f"==========================================")
    t0 = d.tables[0]
    for r_i, r in enumerate(t0.rows):
        for c_i, c in enumerate(r.cells):
            print(f"Cell ({r_i}, {c_i}): paragraphs={len(c.paragraphs)}")
            for p_i, p in enumerate(c.paragraphs):
                print(f"  P{p_i}: {repr(p.text)}")
