import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher

for so in ['SY1615', 'SY1617', 'SY1618', 'SY1609', 'SY1596', 'SY1606']:
    res = sterp_data_fetcher.fetch_sterp_job_data(so)
    if not res or res.get('status') != 'success':
        print(f"Failed to fetch {so}")
        continue
    
    print(f"\n{'='*30} {so} {'='*30}")
    print("Customer:", res.get('customer'))
    print("Items in SO:")
    for i, itm in enumerate(res.get('items', [])):
        print(f"  Line {i+1}: [{itm.get('ItemCode')}] {itm.get('ItemDesc')}")
        det = itm.get('DetailDesc')
        if det:
            print(f"    Detail: {det[:100]}...")
            
    # Check overall extraction
    all_blocks = []
    all_components = []
    for itm in res['items']:
        all_blocks.append(f"{itm['ItemDesc']} {itm['AddDesc']} {itm['DetailDesc']}")
        all_components.extend(itm.get('bom_components', []))
    for w in res.get('work_orders', []):
        for sub in w.get('subitems', []):
            all_blocks.append(f"{sub['ItemDesc']} {sub['ModelNo']} {sub['DetailDesc']}")
            all_components.append(sub)
    for jc in res.get('job_cards', []):
        all_blocks.append(f"{jc['ItemDesc']} {jc['ModelNo']} {jc['DetailDesc']} {jc['Remark']}")
        all_components.append(jc)
        
    combined = '\n'.join(all_blocks)
    overall = sterp_data_fetcher._extract_specs_from_text(combined, default_job=res['so_no'], customer_name=res['customer'])
    overall = sterp_data_fetcher._enrich_specs_from_components(overall, all_components)
    
    print("Combined Overall Specs:")
    for k in ['capacity', 'indicator_model', 'sensor_model', 'load_cell_model', 'lc_model', 'lc_capacity', 'lc_qty', 'jbox_model', 'jbox_model1', 'voltage', 'tracking_no']:
        if overall.get(k):
            print(f"  {k}: {overall.get(k)}")
