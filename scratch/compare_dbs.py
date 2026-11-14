import sqlite3, os

p_local = r"C:\Users\Admin\OneDrive\Desktop\PDN TO QC REGISTER\database\register.db"
p_net = r"\\192.168.100.248\prdndata\ABEL\SFT\PDN TO QC REGISTER\database\register.db"

conn_loc = sqlite3.connect(p_local)
conn_net = sqlite3.connect(p_net)

loc_row = conn_loc.execute("SELECT MAX(CAST(sl_no AS INTEGER)) FROM register WHERE sheet_type='2026' AND ((job_no IS NOT NULL AND TRIM(job_no) != '') OR (customer IS NOT NULL AND TRIM(customer) != ''))").fetchone()
net_row = conn_net.execute("SELECT MAX(CAST(sl_no AS INTEGER)) FROM register WHERE sheet_type='2026' AND ((job_no IS NOT NULL AND TRIM(job_no) != '') OR (customer IS NOT NULL AND TRIM(customer) != ''))").fetchone()

print("Local max filled sl_no:", loc_row[0])
print("Network max filled sl_no:", net_row[0])

# Check latest 3 jobs in network DB
latest_net = conn_net.execute("""
    SELECT id, sl_no, date, job_no, work_order_no, job_card_no, customer, system, items, qty, model_no, serial_no, production_engineer
    FROM register 
    WHERE sheet_type='2026' AND job_no != '' AND job_no IS NOT NULL
    ORDER BY id DESC LIMIT 3
""").fetchall()

print("\nLatest 3 jobs in Network DB:")
for r in latest_net:
    print(r)
