import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")

sql = """
SELECT DISTINCT wod.BOMItem, im.ItemDesc, im.ModelNo, wod.DetailDesc
FROM WorkOrderDetail wod
JOIN ItemMaster im ON wod.BOMItem = im.ItemCode
WHERE (im.ItemCode LIKE 'IJBX%' OR im.ItemDesc LIKE 'JUNCTION BOX%')
"""

cmd = SqlCommand(sql, conn)
r = cmd.ExecuteReader()
print("All Junction Boxes in WorkOrderDetail:")
while r.Read():
    code = str(r[0] or "").strip()
    desc = str(r[1] or "").strip()
    model = str(r[2] or "").strip()
    detail = str(r[3] or "").strip()
    print(f"  Code: {code} | Desc: {desc} | Model: {model} | WODetail: {detail}")
r.Close()

conn.Close()
