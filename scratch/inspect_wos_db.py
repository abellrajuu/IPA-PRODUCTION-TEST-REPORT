import sys, os
sys.path.insert(0, os.path.abspath('.'))
from sterp_data_fetcher import get_connection, execute_query_net

conn = get_connection('I0012627')
conn.Open()
rows = execute_query_net(conn, "SELECT WONo, SONo, SOItem, ItemCode, WOQty, Remarks, TechRemarks, DrawingNo FROM WorkOrderMaster WHERE SONo LIKE '%0192%'", {})
for r in rows:
    print(r)
conn.Close()
