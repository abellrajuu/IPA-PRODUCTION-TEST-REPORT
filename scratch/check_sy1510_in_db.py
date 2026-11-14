import sqlite3

p_net = r"\\192.168.100.248\prdndata\ABEL\SFT\PDN TO QC REGISTER\database\register.db"
conn_net = sqlite3.connect(p_net)
conn_net.row_factory = sqlite3.Row

rows = conn_net.execute("SELECT * FROM register WHERE job_no LIKE '%1510%'").fetchall()
print("Matches for 1510 in Network DB:", len(rows))
for r in rows:
    print(dict(r))
