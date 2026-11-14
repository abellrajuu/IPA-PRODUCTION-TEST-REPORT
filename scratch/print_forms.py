with open(r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\inprocess.html", "r", encoding="utf-8") as f:
    text = f.read()

idx = text.find("<form")
idx_end = text.find("</form>")
print("inprocess.html form content:")
print(text[idx:idx+3000])

with open(r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\inprocess_register.html", "r", encoding="utf-8") as f:
    text2 = f.read()

idx2 = text2.find("<form")
print("\n--- inprocess_register.html form content: ---")
print(text2[idx2:idx2+3000])
