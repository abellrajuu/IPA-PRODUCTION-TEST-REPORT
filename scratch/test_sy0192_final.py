import sys, os
sys.path.insert(0, os.path.abspath('.'))
sys.stdout.reconfigure(encoding='utf-8')
import json
from sterp_data_fetcher import fetch_sterp_job_data

res = fetch_sterp_job_data('SY0192')
print("Status:", res.get("status"))
print("SO:", res.get("so_no"))
print("Customer:", res.get("customer"))
print("Multi Systems Count:", len(res.get("multi_systems", {})))
print("System Names:", res.get("system_names"))

print("\n--- PRIMARY EXTRACTED (FIRST SYSTEM) ---")
ext = res.get("extracted", {})
for k in ["customer", "di_model", "indicator_model", "load_cell_model", "lc_qty", "capacity", "jbox_model", "tag_no"]:
    print(f"  {k}: {ext.get(k)}")

print("\n--- MULTI SYSTEMS BREAKDOWN ---")
for s_name, s_spec in res.get("multi_systems", {}).items():
    tag = s_spec.get("tag_no")
    ind = s_spec.get("indicator_model")
    lc = s_spec.get("load_cell_model")
    cap = s_spec.get("capacity")
    qty = s_spec.get("lc_qty")
    jb = s_spec.get("jbox_model")
    print(f"[{s_name}]")
    print(f"   Tag: {tag} | Indicator: {ind} | LC: {lc} (Cap: {cap}, Qty: {qty}) | JB: {jb}")
