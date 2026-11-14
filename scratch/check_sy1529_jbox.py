import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")

# Find all Work Orders for SY1529
cmd = SqlCommand("SELECT WONo, SONo, ItemCode FROM WorkOrderMaster WHERE SONo LIKE '%1529%' OR WONo LIKE '%1529%'", conn)
r = cmd.ExecuteReader()
wos = []
while r.Read():
    wos.append(str(r["WONo"]))
r.Close()

for wono in wos:
    cmd = SqlCommand(f"SELECT wod.BOMItem, im.ItemDesc, im.ModelNo, im.DetailDesc, wod.DetailDesc as WODetail FROM WorkOrderDetail wod LEFT JOIN ItemMaster im ON wod.BOMItem = im.ItemCode WHERE wod.WONo='{wono}'", conn)
    r = cmd.ExecuteReader()
    while r.Read():
        desc = str(r["ItemDesc"] or "")
        model = str(r["ModelNo"] or "")
        detail = str(r["DetailDesc"] or "")
        wodetail = str(r["WODetail"] or "")
        bom_item = str(r["BOMItem"] or "")
        all_text = f"{desc} {model} {detail} {wodetail}"
        if any(k in all_text.lower() for k in ["junction", "jbox", "box"]):
            print(f"[{wono}] BOM: {bom_item} | Desc: {desc} | Model: {model} | Detail: {detail} | WODetail: {wodetail}")
    r.Close()

conn.Close()
