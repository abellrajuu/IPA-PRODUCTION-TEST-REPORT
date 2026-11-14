import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")

# Find where junction box is mentioned in Work Orders or BOM or Sales Orders
queries = [
    # 1. WorkOrderDetail joined with ItemMaster
    """
    SELECT TOP 20 wod.WONo, wod.ItemCode, im.ItemDesc, im.DetailDesc, im.ModelNo, wod.DetailDesc as WODetail
    FROM WorkOrderDetail wod
    LEFT JOIN ItemMaster im ON wod.BOMItem = im.ItemCode
    WHERE im.ItemDesc LIKE '%JUNCTION%' OR im.ItemCode LIKE 'IJBX%' OR wod.DetailDesc LIKE '%JUNCTION%'
    """,
    # 2. WorkOrderMaster joined with ItemMaster
    """
    SELECT TOP 20 wom.WONo, wom.SONo, wom.ItemCode, im.ItemDesc, im.DetailDesc, im.ModelNo
    FROM WorkOrderMaster wom
    LEFT JOIN ItemMaster im ON wom.ItemCode = im.ItemCode
    WHERE im.ItemDesc LIKE '%JUNCTION%' OR im.ItemCode LIKE 'IJBX%'
    """,
    # 3. ItemMaster where ItemCode LIKE 'IJBX%'
    """
    SELECT TOP 30 ItemCode, ItemDesc, DetailDesc, ModelNo
    FROM ItemMaster
    WHERE ItemCode LIKE 'IJBX%' OR ItemDesc LIKE '%JUNCTION%BOX%'
    """
]

for idx, sql in enumerate(queries):
    print(f"\n=== QUERY {idx+1} ===")
    try:
        cmd = SqlCommand(sql, conn)
        r = cmd.ExecuteReader()
        while r.Read():
            row = [str(r[i] or "").strip() for i in range(r.FieldCount)]
            print(" | ".join(row))
        r.Close()
    except Exception as e:
        print("Error:", e)

conn.Close()
