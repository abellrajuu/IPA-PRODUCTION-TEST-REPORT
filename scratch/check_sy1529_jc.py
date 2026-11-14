import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")
cmd = SqlCommand("SELECT jc.JobCardNo, jc.WONo, jc.ItemCode, im.ItemDesc, im.ModelNo, im.DetailDesc, jc.Remark FROM JobCardMaster jc LEFT JOIN ItemMaster im ON jc.ItemCode = im.ItemCode WHERE jc.WONo IN (SELECT WONo FROM WorkOrderMaster WHERE SONo LIKE '%1529%')", conn)
r = cmd.ExecuteReader()
while r.Read():
    desc = str(r["ItemDesc"] or "")
    model = str(r["ModelNo"] or "")
    detail = str(r["DetailDesc"] or "")
    remark = str(r["Remark"] or "")
    itemcode = str(r["ItemCode"] or "")
    print(f"JC: {r['JobCardNo']} | WONo: {r['WONo']} | Item: {itemcode} | Desc: {desc[:40]} | Model: {model} | Remark: {remark}")
r.Close()
conn.Close()
