import sqlite3, os

db_path = r"C:\Users\Admin\OneDrive\Desktop\PDN TO QC REGISTER\database\register.db"
conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row

tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
print("Tables in register.db:", [t[0] for t in tables])

for t in tables:
    t_name = t[0]
    schema = conn.execute(f"SELECT sql FROM sqlite_master WHERE name='{t_name}'").fetchone()
    print(f"\n--- Schema for {t_name} ---")
    print(schema[0])
    count = conn.execute(f"SELECT count(*) FROM {t_name}").fetchone()[0]
    print(f"Total rows in {t_name}: {count}")
    
    # print the latest 3 rows
    latest = conn.execute(f"SELECT * FROM {t_name} ORDER BY rowid DESC LIMIT 3").fetchall()
    print("Latest 3 rows:")
    for row in latest:
        print(dict(row))
