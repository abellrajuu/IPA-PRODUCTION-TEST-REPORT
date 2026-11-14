import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

conn = sterp_data_fetcher.get_connection("I0012627")

# Find WorkOrder tables
cmd = SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '%WorkOrder%' OR TABLE_NAME LIKE '%JobCard%'", conn)
r = cmd.ExecuteReader()
tables = []
while r.Read():
    tables.append(str(r[0]))
r.Close()
print("Matching tables:", tables)

for tbl in tables:
    print(f"\n--- TABLE: {tbl} ---")
    cmd_col = SqlCommand(f"SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='{tbl}'", conn)
    r_col = cmd_col.ExecuteReader()
    col_names = []
    while r_col.Read():
        col_names.append(str(r_col[0]))
    r_col.Close()
    print("Columns:", col_names)

    # Check rows for SY1529 or 1529
    filter_cols = [c for c in col_names if any(k in c.lower() for k in ['so', 'wo', 'job', 'ref'])]
    where_clauses = [f"{c} LIKE '%1529%'" for c in filter_cols]
    if where_clauses:
        sql = f"SELECT TOP 5 * FROM {tbl} WHERE " + " OR ".join(where_clauses)
        cmd_data = SqlCommand(sql, conn)
        try:
            r_data = cmd_data.ExecuteReader()
            while r_data.Read():
                row_dict = {col_names[i]: str(r_data[i] or "") for i in range(r_data.FieldCount)}
                print(f"  Row in {tbl}:")
                for k, v in row_dict.items():
                    if v.strip():
                        print(f"    {k}: {v}")
            r_data.Close()
        except Exception as e:
            print("  Query error:", e)

conn.Close()
