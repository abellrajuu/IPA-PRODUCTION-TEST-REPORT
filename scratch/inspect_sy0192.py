import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher

res = sterp_data_fetcher.fetch_sterp_job_data('SY0192')
print('Customer:', res.get('customer'))
print('SO No:', res.get('so_no'))
print('PO No:', res.get('raw_so', {}).get('CustPONo'))
print('PO Date:', res.get('raw_so', {}).get('PODate'))
print('SO Date:', res.get('raw_so', {}).get('SODate'))
print('Database:', res.get('database'))
print('Has multi systems:', res.get('has_multiple_systems'))
print('System names:', res.get('system_names'))

print('\n=== ITEMS ===')
for i, itm in enumerate(res.get('items', []), start=1):
    print(f"[{i}] ItemCode: {itm.get('ItemCode')} | Desc: {itm.get('ItemDesc')} | Qty: {itm.get('Quantity')}")
    print(f"    Detail: {repr(itm.get('DetailDesc'))}")

print('\n=== WORK ORDERS ===')
for wo in res.get('work_orders', []):
    print(f"WO: {wo.get('WONo')} | ItemCode: {wo.get('ItemCode')} | BOM: {wo.get('BOMNo')} | WOQty: {wo.get('WOQty')}")
    for sub in wo.get('subitems', []):
        desc = sub.get('ItemDesc', '')
        detail = sub.get('DetailDesc', '')
        if any(k in (desc + detail).upper() for k in ['LOADCELL', 'INDICATOR', 'JUNCTION', 'PANEL', 'DISPLAY', 'MOUNTING', 'BOX']):
            print(f"   Sub: {sub.get('BOMItem')} | {desc} | Model: {sub.get('ModelNo')} | Detail: {detail[:70]}")

print('\n=== SYSTEM 1 to 6 EXTRACTED SPECS ===')
multi = res.get('multi_systems', {})
for s_name in res.get('system_names', []):
    s = multi.get(s_name, {})
    print(f"\n--- {s_name} ---")
    print(f"  Capacity: {s.get('capacity')}")
    print(f"  Tag/Bin No: {s.get('tag_no')}")
    print(f"  Indicator: {s.get('di_model')}")
    print(f"  Load Cell 1: {s.get('lc_model1')} ({s.get('lc_capacity1')})")
    print(f"  Load Cell 2: {s.get('lc_model2')} ({s.get('lc_capacity2')})")
    print(f"  Junction Box: {s.get('jbox_model1')}")
    print(f"  Qty: {s.get('lc_qty')}")
