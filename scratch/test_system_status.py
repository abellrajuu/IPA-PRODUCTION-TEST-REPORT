import sys, os, urllib.request, urllib.parse, json, http.cookiejar

cookie_jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))

print("=== RUNNING SYSTEM COMPREHENSIVE HEALTH CHECK ===")

# 1. Login Test
login_data = urllib.parse.urlencode({'username': 'a', 'password': 'a'}).encode('utf-8')
req = urllib.request.Request('http://localhost:5005/login', data=login_data)
try:
    resp = opener.open(req)
    print(f"[PASS] Login (user: 'a'): Status {resp.status}, Redirected to: {resp.geturl()}")
except Exception as e:
    print(f"[FAIL] Login failed: {e}")
    sys.exit(1)

# 2. Portal Test
try:
    resp = opener.open('http://localhost:5005/portal')
    html = resp.read().decode('utf-8')
    assert 'Test Reports' in html and 'View All Reports' in html
    print(f"[PASS] Portal: Status {resp.status}, Original 2-card workstation UI verified.")
except Exception as e:
    print(f"[FAIL] Portal check failed: {e}")

# 3. Digital Indicator Route Test
for route in ['/digital-indicator', '/digital_indicator']:
    try:
        resp = opener.open(f'http://localhost:5005{route}')
        html = resp.read().decode('utf-8')
        assert 'Digital Indicator Test Report Form' in html
        print(f"[PASS] Route '{route}': Status {resp.status}, Form loaded properly.")
    except Exception as e:
        print(f"[FAIL] Route '{route}' failed: {e}")

# 4. Multi-system Job Fetch Test (SY0192)
try:
    resp = opener.open('http://localhost:5005/api/fetch-job-data?job_no=SY0192')
    data = json.loads(resp.read().decode('utf-8'))
    assert data.get('status') == 'success'
    has_ms = data.get('has_multiple_systems')
    n_sys = len(data.get('multi_systems', []))
    print(f"[PASS] ST-ERP Fetch (SY0192): Status {resp.status}, has_multiple_systems={has_ms}, systems_count={n_sys}")
except Exception as e:
    print(f"[FAIL] SY0192 fetch failed: {e}")

# 5. Review Dashboard Test
try:
    resp = opener.open('http://localhost:5005/review-dashboard')
    html = resp.read().decode('utf-8')
    assert 'SAVED TEST REPORTS' in html
    print(f"[PASS] Review Dashboard: Status {resp.status}, Original clean table verified.")
except Exception as e:
    print(f"[FAIL] Review Dashboard check failed: {e}")

# 6. API Pending Reviews
try:
    resp = opener.open('http://localhost:5005/api/pending-reviews')
    data = json.loads(resp.read().decode('utf-8'))
    print(f"[PASS] API Pending Reviews: Status {resp.status}, Found {len(data.get('records', []))} records.")
except Exception as e:
    print(f"[FAIL] API Pending Reviews failed: {e}")

print("=== ALL HEALTH CHECKS COMPLETE ===")
