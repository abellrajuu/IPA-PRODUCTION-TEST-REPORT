import sys, os
sys.path.insert(0, os.path.abspath('.'))
from sterp_data_fetcher import get_connection
import clr
clr.AddReference('System.Data')
from System.Data.SqlClient import SqlCommand

conn = get_connection('I0012627')
cmd = SqlCommand("SELECT WONo, SONo, ItemCode, WOQty, Remark, DSNo, MWONo FROM WorkOrderMaster WHERE SONo LIKE '%0192%'", conn)
r = cmd.ExecuteReader()
while r.Read():
    print(f"WO: {r.GetValue(0)} | Qty: {r.GetValue(3)} | Remark: {r.GetValue(4)} | DSNo: {r.GetValue(5)} | MWONo: {r.GetValue(6)}")
conn.Close()
