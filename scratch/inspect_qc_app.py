with open(r"C:\Users\Admin\OneDrive\Desktop\PDN TO QC REGISTER\app.py", "r", encoding="utf-8") as f:
    text = f.read()

import re
matches = [l for l in text.splitlines() if "def " in l]
print("Functions in PDN TO QC REGISTER app.py:")
for m in matches[:35]:
    print(" ", m)
