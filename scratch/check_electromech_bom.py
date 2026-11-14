import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection('I0012627')
cmd = SqlCommand("SELECT TOP 20 bd.BOMNo, bd.ACode, im.ItemDesc, im.ModelNo, bd.DetailDesc FROM BOMDetails bd LEFT JOIN ItemMaster im ON bd.ACode = im.ItemCode WHERE bd.BOMNo IN (SELECT BOMNo FROM BOMMaster WHERE ItemCode = 'ISYS00000WA9')", conn)
r = cmd.ExecuteReader()
print('BOM items for ISYS00000WA9:')
while r.Read():
    print(f"  {r.GetValue(0)} | {r.GetValue(1)} | {r.GetValue(2)} | {r.GetValue(3)} | {str(r.GetValue(4))[:60]}")
r.Close()

# Also check other Electromech jobs from 2026/2025 with complete WOs
cmd2 = SqlCommand("SELECT TOP 10 so.SONo, so.SODate, so.SPartyName, wo.WONo, wo.ItemCode, im.ItemDesc FROM SalesOrderMaster so INNER JOIN WorkOrderMaster wo ON so.SONo = wo.SONo LEFT JOIN ItemMaster im ON wo.ItemCode = im.ItemCode WHERE UPPER(so.SPartyName) LIKE '%ELECTROMECH%' ORDER BY so.SODate DESC", conn)
r2 = cmd2.ExecuteReader()
print('\nElectromech jobs with Work Orders:')
while r2.Read():
    print(f"  SO: {r2.GetValue(0)} | Date: {r2.GetValue(1)} | WO: {r2.GetValue(3)} | Item: {r2.GetValue(4)} | {r2.GetValue(5)}")
r2.Close()
conn.Close()
