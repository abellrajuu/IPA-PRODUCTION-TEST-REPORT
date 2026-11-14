import docx

def compare_docs(og_path, gen_path):
    print(f"\n=======================================================")
    print(f"COMPARING:\n  OG:  {og_path}\n  GEN: {gen_path}")
    print(f"=======================================================")
    d_og = docx.Document(og_path)
    d_gen = docx.Document(gen_path)

    # 1. Paragraphs count and differences
    print(f"Paragraphs count: OG={len(d_og.paragraphs)}, GEN={len(d_gen.paragraphs)}")
    for i in range(min(len(d_og.paragraphs), len(d_gen.paragraphs))):
        p_og = d_og.paragraphs[i]
        p_gen = d_gen.paragraphs[i]
        t_og = p_og.text.strip()
        t_gen = p_gen.text.strip()
        if t_og != t_gen:
            print(f"  P{i} DIFF:")
            print(f"    OG:  '{t_og}'")
            print(f"    GEN: '{t_gen}'")
            # font size
            og_sizes = [r.font.size.pt if r.font.size else None for r in p_og.runs]
            gen_sizes = [r.font.size.pt if r.font.size else None for r in p_gen.runs]
            print(f"    OG fonts:  {og_sizes}")
            print(f"    GEN fonts: {gen_sizes}")

    # 2. Tables count and structure
    print(f"\nTables count: OG={len(d_og.tables)}, GEN={len(d_gen.tables)}")
    for t_i in range(min(len(d_og.tables), len(d_gen.tables))):
        t_og = d_og.tables[t_i]
        t_gen = d_gen.tables[t_i]
        print(f"  Table {t_i}: OG rows={len(t_og.rows)}, cols={len(t_og.columns)} | GEN rows={len(t_gen.rows)}, cols={len(t_gen.columns)}")
        # Check first row cell texts
        for r_i in range(min(len(t_og.rows), len(t_gen.rows))):
            r_og = t_og.rows[r_i]
            r_gen = t_gen.rows[r_i]
            og_texts = [c.text.replace('\n', ' ').strip() for c in r_og.cells]
            gen_texts = [c.text.replace('\n', ' ').strip() for c in r_gen.cells]
            if r_i == 0:
                print(f"    Row 0 OG:  {og_texts}")
                print(f"    Row 0 GEN: {gen_texts}")
            # check cell fonts
            for c_i in range(min(len(r_og.cells), len(r_gen.cells))):
                c_og = r_og.cells[c_i]
                c_gen = r_gen.cells[c_i]
                og_f = [r.font.size.pt if r.font.size else None for p in c_og.paragraphs for r in p.runs]
                gen_f = [r.font.size.pt if r.font.size else None for p in c_gen.paragraphs for r in p.runs]
                if og_f and gen_f and og_f != gen_f and r_i == 0:
                    print(f"    Cell (r{r_i}, c{c_i}) fonts: OG={og_f} vs GEN={gen_f}")

compare_docs("DOC/PP-05_06 Digital Indicator.docx", "tests/temp_variety_outputs/Digital_Indicator_4LC_Load_1AO_Report.docx")
compare_docs("DOC/Belt Scale.docx", "tests/temp_variety_outputs/Belt_Scale_4LC_1AO_Report.docx")
