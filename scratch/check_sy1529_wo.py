import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")

# Check IJBX00000419 in ItemMaster, BOMMaster, BOMDetails
for tbl in ["ItemMaster", "BOMMaster", "BOMDetails"]:
    cmd = SqlCommand(f"SELECT * FROM {tbl} WHERE ItemCode='IJBX00000419' OR BOMNo='IJBX00000419'", conn)
    try:
        r = cmd.ExecuteReader()
        cols = [r.GetName(i) for i in range(r.FieldCount)]
        while r.Read():
            print(f"[{tbl}]")
            for c in cols:
                v = str(r[c] or "")
                if v:
                    print(f"  {c}: {v}")
        r.Close()
    except Exception as e:
        print(f"Error {tbl}:", e)

# Also check WorkOrder for SY1529 specifically to see what items are in the Work Orders of SY1529!
print("\n--- Work Orders for SY1529 ---")
cmd = SqlCommand("SELECT WONo, SONo, ItemCode, WOQty, WOStatus FROM WorkOrderMaster WHERE SONo LIKE '%1529%' OR WONo LIKE '%1529%'", conn)
r = cmd.ExecuteReader()
wos = []
while r.Read():
    wos.append(str(r["WONo"]))
    print(f"  WONo: {r['WONo']}, SONo: {r['SONo']}, ItemCode: {r['ItemCode']}")
r.Close()

# For each WONo, find ALL items in WorkOrderDetail
for wono in wos:
    print(f"\n--- Items in WorkOrderDetail for {wono} ---")
    cmd = SqlCommand(f"SELECT wod.*, im.ItemDesc, im.DetailDesc, im.ModelNo FROM WorkOrderDetail wod LEFT JOIN ItemMaster im ON wod.BOMItem = im.ItemCode WHERE wod.WONo='{wono}'", conn)
    r = cmd.ExecuteReader()
    cols = [r.GetName(i) for i in range(r.FieldCount)]
    while r.Read():
        desc = str(r["ItemDesc"] or "")
        model = str(r["ModelNo"] or "")
        detail = str(r["DetailDesc"] or "")
        bom_item = str(r["BOMItem"] or "")
        print(f"  BOMItem: {bom_item} | Desc: {desc} | Model: {model} | Detail: {detail[:100]}")
    r.Close()

conn.Close()
