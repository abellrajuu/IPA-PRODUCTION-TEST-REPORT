import sys, os
sys.path.insert(0, os.path.abspath('.'))
from sterp_data_fetcher import get_connection
import clr
clr.AddReference('System.Data')
from System.Data.SqlClient import SqlCommand

conn = get_connection('I0012627')
cmd = SqlCommand("SELECT AutoNo, WONo, ItemCode, WOQty FROM WorkOrderMaster WHERE SONo LIKE '%0192%'", conn)
r = cmd.ExecuteReader()
while r.Read():
    print(f"WO AutoNo: {r.GetValue(0)} | WONo: {r.GetValue(1)} | ItemCode: {r.GetValue(2)} | Qty: {r.GetValue(3)}")
r.Close()

cmd2 = SqlCommand("SELECT AutoNo, SrNo, ItemCode, Quantity, DetailDesc FROM SalesOrderDetail WHERE SONo LIKE '%0192%'", conn)
r2 = cmd2.ExecuteReader()
while r2.Read():
    print(f"SOD AutoNo: {r2.GetValue(0)} | SrNo: {r2.GetValue(1)} | ItemCode: {r2.GetValue(2)} | Qty: {r2.GetValue(3)} | Desc: {r2.GetValue(4)[:30]}")
r2.Close()
conn.Close()
