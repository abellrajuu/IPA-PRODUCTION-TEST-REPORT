with open(r"C:\Users\Admin\OneDrive\Desktop\PDN TO QC REGISTER\app.py", "r", encoding="utf-8") as f:
    text = f.read()

idx = text.find("def create_job():")
if idx == -1:
    idx = text.find("def add_data():")
print(text[idx:idx+2500])
