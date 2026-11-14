import sys, os

app_js_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_static_assets\js\app.js'
with open(app_js_path, 'r', encoding='utf-8') as f:
    text = f.read()

snippet = """
// ══════════════════════════════════════════════════════════════════════════
// UI-UX PRO MAX: GLOBAL URL PARAMETER AUTO-LAUNCHER
// ══════════════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
    try {
        const urlParams = new URLSearchParams(window.location.search);
        const jobParam = urlParams.get('job_no') || urlParams.get('job');
        if (jobParam) {
            const cleanJob = jobParam.trim().toUpperCase();
            const jobInput = document.querySelector('input[name="job_no"], input[id="job_no"], input[name="job_number"]');
            if (jobInput) {
                jobInput.value = cleanJob;
                setTimeout(() => {
                    if (typeof window.fetchJobNoData === 'function') {
                        window.fetchJobNoData();
                    } else if (typeof fetchJobNoData === 'function') {
                        fetchJobNoData();
                    } else {
                        jobInput.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }, 250);
            }
        }
    } catch(e) {
        console.warn('Job URL auto-launcher notice:', e);
    }
});
"""

if 'UI-UX PRO MAX: GLOBAL URL PARAMETER AUTO-LAUNCHER' not in text:
    text += snippet
    with open(app_js_path, 'w', encoding='utf-8') as f:
        f.write(text)
    print('SUCCESS: Added URL Parameter Auto-Launcher to app.js')
else:
    print('Already present')
