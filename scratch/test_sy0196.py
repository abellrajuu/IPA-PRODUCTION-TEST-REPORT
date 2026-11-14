import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher

# Test fetch_sterp_job_data with the candidate fix
res = sterp_data_fetcher.fetch_sterp_job_data("SY0196")
print("Status:", res.get("status"))
print("Database:", res.get("database"))
print("Customer:", res.get("customer"))
print("Items:", len(res.get("items", [])))
print("Work Orders:", len(res.get("work_orders", [])))
print("Extracted:")
for k, v in res.get("extracted", {}).items():
    if v:
        print(f"  {k}: {v}")
