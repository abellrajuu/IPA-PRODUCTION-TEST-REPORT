import docx

for doc_path in ["DOC/Belt Scale.docx", "DOC/PP-05_06 Digital Indicator.docx"]:
    doc = docx.Document(doc_path)
    sec = doc.sections[0]
    print(f"\n--- {doc_path} ---")
    print(f"Page width: {sec.page_width.pt} pt ({sec.page_width.inches} in, {int(sec.page_width.inches * 1440)} dxa)")
    print(f"Page height: {sec.page_height.pt} pt ({sec.page_height.inches} in, {int(sec.page_height.inches * 1440)} dxa)")
    print(f"Left margin: {sec.left_margin.pt} pt ({sec.left_margin.inches} in, {int(sec.left_margin.inches * 1440)} dxa)")
    print(f"Right margin: {sec.right_margin.pt} pt ({sec.right_margin.inches} in, {int(sec.right_margin.inches * 1440)} dxa)")
    print(f"Top margin: {sec.top_margin.pt} pt ({sec.top_margin.inches} in)")
    print(f"Bottom margin: {sec.bottom_margin.pt} pt ({sec.bottom_margin.inches} in)")
    usable_width = sec.page_width.pt - sec.left_margin.pt - sec.right_margin.pt
    print(f"Usable width: {usable_width} pt ({usable_width / 72.0} in, {int(usable_width * 20)} dxa)")
