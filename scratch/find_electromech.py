import sqlite3, os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import sterp_data_fetcher

dbs = sterp_data_fetcher.get_available_year_databases()
print("Found databases:", len(dbs))

electromech_jobs = []

for db_info in dbs:
    db_path = db_info.get("path") if isinstance(db_info, dict) else db_info
    db_name = db_info.get("name") if isinstance(db_info, dict) else os.path.basename(db_path)
    if not os.path.exists(db_path):
        continue
    try:
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        
        # Check tables in this DB
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [r[0] for r in cur.fetchall()]
        
        # Check sales order tables
        for t in tables:
            cur.execute(f"PRAGMA table_info({t})")
            cols = [c[1] for c in cur.fetchall()]
            
            # Find customer / party / client columns
            cust_cols = [c for c in cols if any(k in c.lower() for k in ['cust', 'party', 'client', 'name', 'ledger'])]
            job_cols = [c for c in cols if any(k in c.lower() for k in ['sono', 'so_no', 'job', 'docno', 'doc_no', 'voucherno', 'orderno'])]
            
            if cust_cols and job_cols:
                for cc in cust_cols:
                    for jc in job_cols:
                        try:
                            cur.execute(f"SELECT DISTINCT {jc}, {cc} FROM {t} WHERE UPPER({cc}) LIKE '%ELECTROMECH%' LIMIT 20")
                            rows = cur.fetchall()
                            for r in rows:
                                if r[0]:
                                    electromech_jobs.append({
                                        "db": db_name,
                                        "table": t,
                                        "job_no": str(r[0]).strip(),
                                        "customer": str(r[1]).strip()
                                    })
                        except Exception:
                            pass
        conn.close()
    except Exception as e:
        print(f"Error {db_name}: {e}")

# Deduplicate
seen = set()
unique_jobs = []
for item in electromech_jobs:
    k = (item['db'], item['job_no'])
    if k not in seen:
        seen.add(k)
        unique_jobs.append(item)

print(f"\nTotal Electromech jobs found: {len(unique_jobs)}")
for j in unique_jobs[:30]:
    print(f"DB: {j['db']} | Job No: {j['job_no']} | Customer: {j['customer']}")
