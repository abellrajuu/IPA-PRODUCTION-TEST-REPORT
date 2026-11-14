import sqlite3

db_path = r"C:\Users\Admin\OneDrive\Desktop\PDN TO QC REGISTER\database\register.db"
conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row

# Check distinct sheet_types
types = conn.execute("SELECT DISTINCT sheet_type, count(*) FROM register GROUP BY sheet_type").fetchall()
print("Sheet types:", [dict(t) for t in types])

# Get latest 10 rows with job_no not empty in 2026
latest_2026 = conn.execute("""
    SELECT id, sheet_type, sl_no, date, job_no, work_order_no, job_card_no, customer, system, items, qty, model_no, serial_no, production_engineer
    FROM register 
    WHERE sheet_type='2026' AND job_no != '' AND job_no IS NOT NULL
    ORDER BY id DESC LIMIT 10
""").fetchall()

print("\nLatest 10 populated rows in 2026:")
for r in latest_2026:
    print(dict(r))

# Find the max sl_no or current end of 2026 sheet
max_row = conn.execute("SELECT id, sl_no, job_no, customer FROM register WHERE sheet_type='2026' ORDER BY id DESC LIMIT 5").fetchall()
print("\nTop 5 rows in 2026 sheet:")
for r in max_row:
    print(dict(r))
