import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")

for sono in ['SY1529', 'SY1615', 'SY1605', 'SP1631', 'SP1336', 'SY1294']:
    print(f"\n==========================================")
    print(f"SEARCHING FOR SO: {sono}")
    print(f"==========================================")
    cmd = SqlCommand(f"SELECT wom.WONo, wom.SONo, wom.ItemCode, im.ItemDesc, im.DetailDesc, im.ModelNo FROM WorkOrderMaster wom LEFT JOIN ItemMaster im ON wom.ItemCode=im.ItemCode WHERE wom.SONo='{sono}' OR wom.WONo='{sono}'", conn)
    r = cmd.ExecuteReader()
    wos = []
    while r.Read():
        wos.append(str(r["WONo"]))
        print(f"  WO: {r['WONo']} | Item: {r['ItemCode']} | Desc: {str(r['ItemDesc'])[:40]} | Model: {r['ModelNo']}")
    r.Close()

    for wono in wos:
        # Check subitems / BOM in WorkOrderDetail
        cmd2 = SqlCommand(f"SELECT wod.BOMItem, im.ItemDesc, im.DetailDesc, im.ModelNo, wod.DetailDesc as WODetail FROM WorkOrderDetail wod LEFT JOIN ItemMaster im ON wod.BOMItem=im.ItemCode WHERE wod.WONo='{wono}'", conn)
        r2 = cmd2.ExecuteReader()
        while r2.Read():
            desc = str(r2["ItemDesc"] or "").strip()
            model = str(r2["ModelNo"] or "").strip()
            detail = str(r2["DetailDesc"] or "").strip()
            wodetail = str(r2["WODetail"] or "").strip()
            bom = str(r2["BOMItem"] or "").strip()
            text = f"{desc} {model} {detail} {wodetail}"
            if any(k in text.lower() for k in ["junction", "jbox", "jt", "jc"]):
                print(f"    [WorkOrderDetail] BOM: {bom} | Desc: {desc} | Model: {model} | Detail: {detail} | WODetail: {wodetail}")
        r2.Close()

        # Check JobCardMaster
        cmd3 = SqlCommand(f"SELECT jc.JobCardNo, jc.ItemCode, im.ItemDesc, im.DetailDesc, im.ModelNo, jc.Remark FROM JobCardMaster jc LEFT JOIN ItemMaster im ON jc.ItemCode=im.ItemCode WHERE jc.WONo='{wono}'", conn)
        r3 = cmd3.ExecuteReader()
        while r3.Read():
            desc = str(r3["ItemDesc"] or "").strip()
            model = str(r3["ModelNo"] or "").strip()
            detail = str(r3["DetailDesc"] or "").strip()
            remark = str(r3["Remark"] or "").strip()
            item = str(r3["ItemCode"] or "").strip()
            text = f"{desc} {model} {detail} {remark}"
            if any(k in text.lower() for k in ["junction", "jbox", "jt", "jc"]):
                print(f"    [JobCardMaster] Item: {item} | Desc: {desc} | Model: {model} | Detail: {detail} | Remark: {remark}")
        r3.Close()

conn.Close()
