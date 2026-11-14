import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher
from System.Data.SqlClient import SqlCommand

dbs = sterp_data_fetcher.get_available_year_databases()
print(f"Scanning financial year databases: {dbs}")

all_electromech_orders = []

for db_name in dbs:
    try:
        conn = sterp_data_fetcher.get_connection(db_name)
        try:
            cmd = SqlCommand(
                "SELECT SONo, SODate, CustPONo, PODate, PartyCode, SPartyName, SCity "
                "FROM SalesOrderMaster "
                "WHERE UPPER(SPartyName) LIKE '%ELECTROMECH%' "
                "ORDER BY SODate DESC",
                conn
            )
            r = cmd.ExecuteReader()
            while r.Read():
                all_electromech_orders.append({
                    "db": db_name,
                    "so_no": str(r.GetValue(0)).strip(),
                    "so_date": str(r.GetValue(1)).strip(),
                    "po_no": str(r.GetValue(2)).strip(),
                    "po_date": str(r.GetValue(3)).strip(),
                    "party_code": str(r.GetValue(4)).strip(),
                    "party_name": str(r.GetValue(5)).strip(),
                    "city": str(r.GetValue(6)).strip()
                })
            r.Close()
        finally:
            conn.Close()
    except Exception as e:
        print(f"Error querying {db_name}: {e}")

print(f"\nTotal Sales Orders found for ELECTROMECH: {len(all_electromech_orders)}\n")
for o in all_electromech_orders[:25]:
    print(f"[{o['db']}] SO: {o['so_no']:<10} | Date: {o['so_date']:<12} | Customer: {o['party_name']}")

# Now let's fetch full job data using sterp_data_fetcher.fetch_sterp_job_data on the top 3 Electromech jobs
print("\n" + "="*80)
print("TESTING sterp_data_fetcher.fetch_sterp_job_data ON TOP ELECTROMECH JOBS:")
print("="*80)

for o in all_electromech_orders[:4]:
    so = o['so_no']
    print(f"\n>>> FETCHING JOB: {so} <<<")
    data = sterp_data_fetcher.fetch_sterp_job_data(so)
    if data:
        ext = data.get("extracted", {})
        print(f"Customer Name: {ext.get('customer_name')}")
        print(f"Capacity:      {ext.get('capacity')}")
        print(f"JBox Model:    {ext.get('jbox_model')}")
        print(f"JBox Model 1:  {ext.get('jbox_model1')}")
        print(f"JBox Model 2:  {ext.get('jbox_model2')}")
        print(f"Loadcell Model:{ext.get('loadcell_model')}")
        print(f"Speed Sensor:  {ext.get('speed_sensor_model')}")
        print(f"Indicator:     {ext.get('indicator_model')}")
        print(f"Total Systems: {len(ext.get('systems', []))}")
        for i, s in enumerate(ext.get('systems', [])):
            print(f"  System {i+1}: {s.get('system_name')} | Capacity: {s.get('capacity')} | JBox: {s.get('jbox_model')} | LC: {s.get('loadcell_model')} | Indicator: {s.get('indicator_model')}")
        print(f"Work Orders:   {[wo.get('wo_no') for wo in data.get('work_orders', [])]}")
    else:
        print(f"Failed to fetch {so}")
