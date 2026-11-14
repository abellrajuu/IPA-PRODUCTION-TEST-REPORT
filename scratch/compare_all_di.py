import docx

di_pairs = [
    ("DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 4 LC Load.docx"),
    ("DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM with 2 current output.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 4 LC Load - 2 AO.docx"),
    ("DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM mv calibration.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 4 LC mV.docx"),
    ("DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM mv calibration with 2 current output.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 4 LC mV - 2 AO.docx"),
    ("DOC/DIGITAL INDICATOR 8  LOAD CELL SYSTEM.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 8 LC Load.docx"),
    ("DOC/DIGITAL INDICATOR 8  LOAD CELL SYSTEM with 2 current output.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 8 LC Load - 2 AO.docx"),
    ("DOC/DIGITAL INDICATOR 8  LOAD CELL SYSTEM  mv  calibration.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 8 LC mV.docx"),
    ("DOC/DIGITAL INDICATOR 8  LOAD CELL SYSTEM  mv  calibration with 2 current output.docx", "backend_docx_templates/PP-05_06 Digital Indicator - 8 LC mV - 2 AO.docx"),
]

for og, be in di_pairs:
    d1 = docx.Document(og)
    d2 = docx.Document(be)
    print(f"\nComparing {og} vs {be}")
    print(f"  OG paras={len(d1.paragraphs)}, tables={len(d1.tables)} | BE paras={len(d2.paragraphs)}, tables={len(d2.tables)}")
    for t_i in range(len(d1.tables)):
        t1 = d1.tables[t_i]
        t2 = d2.tables[t_i]
        if (len(t1.rows), len(t1.columns)) != (len(t2.rows), len(t2.columns)):
            print(f"    Table {t_i} MISMATCH: OG={len(t1.rows)}x{len(t1.columns)} vs BE={len(t2.rows)}x{len(t2.columns)}")
        else:
            print(f"    Table {t_i} MATCHES: {len(t1.rows)}x{len(t1.columns)}")
