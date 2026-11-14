import sys, os
sys.path.insert(0, os.path.abspath('.'))
sys.stdout.reconfigure(encoding='utf-8')
import json
from sterp_data_fetcher import fetch_sterp_job_data

res = fetch_sterp_job_data('SY0192')
items = res.get('items', [])
print(f"Items count: {len(items)}")
for idx, itm in enumerate(items):
    code = itm.get('ItemCode')
    desc = itm.get('ItemDesc')
    qty = itm.get('Quantity')
    print(f"Item {idx+1}: {code} - {desc} (Qty: {qty})")
    boms = itm.get('bom_components', [])
    print(f"   BOM components count: {len(boms)}")
    for b in boms:
        bcode = b.get('BOMItem')
        bdesc = b.get('ItemDesc')
        bmodel = b.get('ModelNo')
        bdetail = (b.get('DetailDesc') or '')[:50]
        breqqty = b.get('ReqQty')
        print(f"     [{bcode}] {bdesc} | Model: {bmodel} | Detail: {bdetail} | ReqQty: {breqqty}")

wos = res.get('work_orders', [])
print(f"WOs count: {len(wos)}")
for w in wos:
    subs = w.get('subitems', [])
    wono = w.get('WONo')
    print(f"WO {wono} (Qty: {w.get('WOQty')}): {len(subs)} subitems")
    for s in subs:
        scode = s.get('ItemCode')
        sdesc = s.get('ItemDesc')
        smodel = s.get('ModelNo')
        sqty = s.get('ReqQty')
        print(f"   [{scode}] {sdesc} | Model: {smodel} | Qty: {sqty}")
