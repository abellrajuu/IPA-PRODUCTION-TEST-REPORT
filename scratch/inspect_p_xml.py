import docx

doc = docx.Document("DOC/PP-05_06 Digital Indicator.docx")
for idx in [2, 3, 4, 9, 11]:
    p = doc.paragraphs[idx]
    print(f"\n--- Paragraph {idx} ---")
    print("Raw text:", repr(p.text))
    print("XML:", p._p.xml)
