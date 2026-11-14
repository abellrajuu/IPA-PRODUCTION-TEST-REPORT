import sys, os
sys.path.insert(0, os.path.abspath('.'))
sys.stdout.reconfigure(encoding='utf-8')
from sterp_data_fetcher import fetch_sterp_job_data

for j in ['SY1605', 'sy0196', 'SY0192']:
    res = fetch_sterp_job_data(j)
    db = res.get('database')
    cust = res.get('customer')
    ext = res.get('extracted', {})
    ind = ext.get('indicator_model')
    lc = ext.get('load_cell_model')
    qty = ext.get('lc_qty')
    jb = ext.get('jbox_model')
    print(f"=== Job: {j} (DB: {db}) ===")
    print(f"  Customer: {cust}")
    print(f"  Multi Systems: {len(res.get('multi_systems', {}))}")
    print(f"  Indicator: {ind} | LC: {lc} | LC Qty: {qty} | JB: {jb}")
