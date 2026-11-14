import docx, os

b_path = "backend_docx_templates/PP-05_06 Digital Indicator - 4 LC Load.docx"
d_path = "DOC/DIGITAL INDICATOR 4 LOAD CELL SYSTEM.docx"

d_b = docx.Document(b_path)
d_d = docx.Document(d_path)

print(f"backend paras: {len(d_b.paragraphs)}, tables: {len(d_b.tables)}")
print(f"DOC paras: {len(d_d.paragraphs)}, tables: {len(d_d.tables)}")

print("\n--- Backend Tables ---")
for i, t in enumerate(d_b.tables):
    print(f"Table {i}: {len(t.rows)}x{len(t.columns)} | R0: {[c.text.replace(chr(10), ' ')[:25] for c in t.rows[0].cells]}")

print("\n--- DOC Tables ---")
for i, t in enumerate(d_d.tables):
    print(f"Table {i}: {len(t.rows)}x{len(t.columns)} | R0: {[c.text.replace(chr(10), ' ')[:25] for c in t.rows[0].cells]}")
