import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")
cmd = SqlCommand("SELECT TOP 20 so.SONo, p.PartyName FROM SalesOrderMaster so LEFT JOIN PartyMaster p ON so.PartyCode = p.PartyCode ORDER BY so.SODate DESC", conn)
r = cmd.ExecuteReader()
while r.Read():
    sono = str(r["SONo"])
    cust = str(r["PartyName"])
    print(f"{sono}: len={len(cust)} | {cust}")
r.Close()
conn.Close()
