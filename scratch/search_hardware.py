import sys, os
sys.path.insert(0, os.path.abspath('.'))
sys.stdout.reconfigure(encoding='utf-8')
from sterp_data_fetcher import fetch_sterp_job_data

res = fetch_sterp_job_data('SY0192')
wos = res.get('work_orders', [])
for w in wos:
    print(f"=== WO {w.get('WONo')} ({w.get('ItemDesc')}) ===")
    for s in w.get('subitems', []):
        desc = (s.get('ItemDesc') or '')
        detail = (s.get('DetailDesc') or '')
        model = (s.get('ModelNo') or '')
        code = (s.get('ItemCode') or '')
        full = f"{code} | {desc} | {model} | {detail}"
        if any(k in full.upper() for k in ['INDICATOR', 'DS105', 'LOADCELL', 'JUNCTION BOX', 'PANEL', 'SMPS', 'BARRIER']):
            print(f"  -> {code} | {desc} | Model: {model} | Detail: {detail} (Qty: {s.get('ReqQty')})")
