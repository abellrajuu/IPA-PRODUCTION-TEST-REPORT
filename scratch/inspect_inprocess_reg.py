import os, re
import docx

p_html = r"C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\inprocess_register.html"
with open(p_html, "r", encoding="utf-8", errors="ignore") as f:
    text = f.read()

print("=== inprocess_register.html ===")
inputs = re.findall(r'<input[^>]*name=["\']([^"\']+)["\']', text)
print("Input names:", inputs)

for l in text.splitlines():
    if '<th' in l or '<table' in l or 'form' in l:
        print(l.strip()[:100])
