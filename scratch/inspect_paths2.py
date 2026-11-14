with open(r"C:\Users\Admin\OneDrive\Desktop\PDN TO QC REGISTER\app.py", "r", encoding="utf-8") as f:
    text = f.read()

idx = text.find("def resolve_paths")
print(text[idx+1200:idx+2600])
