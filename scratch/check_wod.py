import sys, os
sys.path.insert(0, os.path.abspath('.'))
from sterp_data_fetcher import get_connection
import clr
clr.AddReference('System.Data')
from System.Data.SqlClient import SqlCommand

conn = get_connection('I0012627')
cmd = SqlCommand("SELECT wod.BOMItem, wod.ReqQty, im.ItemDesc, wod.DetailDesc FROM WorkOrderDetail wod LEFT JOIN ItemMaster im ON wod.BOMItem=im.ItemCode WHERE wod.WONo='WY00001458'", conn)
r = cmd.ExecuteReader()
while r.Read():
    code = r.GetValue(0)
    qty = r.GetValue(1)
    desc = r.GetValue(2)
    detail = r.GetValue(3)
    if 'LOADCELL' in str(desc).upper():
        print(f"[{code}] Qty: {qty} | Desc: {desc} | Detail: {detail}")
conn.Close()
