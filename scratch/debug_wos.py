import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")
target_so = "SY1529"

cmd_wos = SqlCommand(
    "SELECT wo.WONo, wo.ItemCode, im.ItemDesc, wo.WOQty, wo.BOMNo "
    "FROM WorkOrderMaster wo "
    "LEFT JOIN ItemMaster im ON wo.ItemCode = im.ItemCode "
    "WHERE wo.SONo=@so",
    conn
)
cmd_wos.Parameters.AddWithValue("@so", target_so)
r_w = cmd_wos.ExecuteReader()
wos = []
while r_w.Read():
    wos.append({
        "WONo": str(r_w.GetValue(0)),
        "ItemCode": str(r_w.GetValue(1)),
        "ItemDesc": str(r_w.GetValue(2)),
        "WOQty": str(r_w.GetValue(3)),
        "BOMNo": str(r_w.GetValue(4)),
        "subitems": []
    })
r_w.Close()
print(f"Found {len(wos)} WOs for {target_so}")

for w in wos:
    print(w)
    cmd_subs = SqlCommand(
        "SELECT wod.BOMItem, im.ItemDesc, im.ModelNo, wod.DetailDesc, wod.ReqQty, im.DrawingNo "
        "FROM WorkOrderDetail wod "
        "LEFT JOIN ItemMaster im ON wod.BOMItem = im.ItemCode "
        "WHERE wod.WONo=@wono",
        conn
    )
    cmd_subs.Parameters.AddWithValue("@wono", w["WONo"])
    r_sub = cmd_subs.ExecuteReader()
    while r_sub.Read():
        item = {
            "BOMItem": str(r_sub.GetValue(0)),
            "ItemDesc": str(r_sub.GetValue(1)),
            "ModelNo": str(r_sub.GetValue(2)),
            "DetailDesc": str(r_sub.GetValue(3)),
            "ReqQty": str(r_sub.GetValue(4)),
            "DrawingNo": str(r_sub.GetValue(5)),
        }
        if any(k in f"{item['ItemDesc']} {item['DetailDesc']}".lower() for k in ["box", "junction", "jbox"]):
            print("  JBOX SUBITEM:", item)
            # Test normalize
            from sterp_data_fetcher import _normalize_jbox_model
            norm = _normalize_jbox_model(item["ItemDesc"], item["DetailDesc"], item["ModelNo"], item["BOMItem"])
            print("  NORM RESULT:", norm)
    r_sub.Close()

conn.Close()
