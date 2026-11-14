import sys, os

di_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\digital_indicator.html'
with open(di_path, 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Upgrade Job No input group
old_job_group = '<div class="input-group"><label>Job No.</label><input type="text" name="job_no" id="job_no"></div>'
new_job_group = """<div class="input-group">
    <label style="display: flex; justify-content: space-between; align-items: center;">
        <span>Job No.</span>
        <span style="font-size: 11px; font-weight: 700; color: #0284c7; background: #e0f2fe; padding: 2px 8px; border-radius: 999px;"><i class="fa-solid fa-database"></i> ST-ERP</span>
    </label>
    <div style="display: flex; gap: 8px;">
        <input type="text" name="job_no" id="job_no" placeholder="e.g. SY0192, SY1246" style="flex: 1;">
        <button type="button" id="btnFetchJobData" onclick="if(typeof window.fetchJobNoData==='function'){window.fetchJobNoData();}else if(typeof fetchJobNoData==='function'){fetchJobNoData();}" style="background: linear-gradient(135deg, #0284c7, #0369a1); color: #ffffff; border: none; padding: 0 16px; border-radius: 8px; font-weight: 700; font-size: 13px; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; box-shadow: 0 2px 8px rgba(2, 132, 199, 0.28); transition: all 0.2s ease;">
            <i class="fa-solid fa-cloud-arrow-down"></i> Fetch
        </button>
    </div>
</div>"""

if old_job_group in text:
    text = text.replace(old_job_group, new_job_group)
    print('SUCCESS: Upgraded Job No input group')
else:
    print('old_job_group not matched exactly, checking variants...')

# 2. Add inputmode="decimal" to table-input elements
text = text.replace('class="table-input"', 'class="table-input" inputmode="decimal"')
print('SUCCESS: Added inputmode="decimal" to matrix cells')

# 3. Clean up the duplicate fetchExcelData script block
# Replace the fetchExcelData block with a clean forwarder to fetchJobNoData
idx1 = text.find('async function fetchExcelData()')
if idx1 != -1:
    script_start = text.rfind('<script>', 0, idx1)
    script_end = text.find('</script>', idx1) + 9
    clean_forwarder = """<script>
    // Unified forwarder to authoritative fetchJobNoData
    async function fetchExcelData() {
        if (typeof window.fetchJobNoData === 'function') return window.fetchJobNoData();
        if (typeof fetchJobNoData === 'function') return fetchJobNoData();
    }
</script>"""
    text = text[:script_start] + clean_forwarder + text[script_end:]
    print('SUCCESS: Replaced redundant fetchExcelData block with clean forwarder')

# 4. Expose window.fetchJobNoData globally in fetchJobNoData block
idx2 = text.find('async function fetchJobNoData()')
if idx2 != -1 and 'window.fetchJobNoData = fetchJobNoData;' not in text:
    text = text.replace('async function fetchJobNoData() {', 'window.fetchJobNoData = async function fetchJobNoData() {')
    print('SUCCESS: Exposed window.fetchJobNoData globally')

with open(di_path, 'w', encoding='utf-8') as f:
    f.write(text)

print('SUCCESS: Fully updated digital_indicator.html with UI-UX Pro Max enhancements!')
