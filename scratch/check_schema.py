import sys, os
sys.path.insert(0, os.path.abspath('.'))
from sterp_data_fetcher import get_connection
import clr
clr.AddReference('System.Data')
from System.Data.SqlClient import SqlCommand

conn = get_connection('I0012627')
cmd = SqlCommand("SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='WorkOrderMaster'", conn)
r = cmd.ExecuteReader()
while r.Read():
    print(r.GetString(0), r.GetString(1))
conn.Close()
