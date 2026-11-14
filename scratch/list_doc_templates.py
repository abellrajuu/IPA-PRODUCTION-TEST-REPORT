import docx, os

doc_files = [f for f in os.listdir("DOC") if f.endswith(".docx") and not f.startswith("~$")]
print(f"Total templates in DOC: {len(doc_files)}")
for f in sorted(doc_files):
    p = os.path.join("DOC", f)
    d = docx.Document(p)
    title = ""
    for para in d.paragraphs[:5]:
        if para.text.strip():
            title = para.text.strip()[:60]
            break
    print(f"  {f:65} | Tables: {len(d.tables)} | Paras: {len(d.paragraphs):2} | Title: {title}")
