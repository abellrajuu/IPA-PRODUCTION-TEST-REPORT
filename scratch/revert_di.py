import sys, os

di_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\digital_indicator.html'
with open(di_path, 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Restore Job No group
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

og_job_group = '<div class="input-group"><label>Job No.</label><input type="text" name="job_no" id="job_no"></div>'

if new_job_group in text:
    text = text.replace(new_job_group, og_job_group)
    print('SUCCESS: Restored Job No input in digital_indicator.html')

# 2. Remove inputmode="decimal"
text = text.replace('class="table-input" inputmode="decimal"', 'class="table-input"')

# 3. Restore fetchExcelData block
clean_forwarder = """<script>
    // Unified forwarder to authoritative fetchJobNoData
    async function fetchExcelData() {
        if (typeof window.fetchJobNoData === 'function') return window.fetchJobNoData();
        if (typeof fetchJobNoData === 'function') return fetchJobNoData();
    }
</script>"""

og_fetch_excel = """<script>
    async function fetchExcelData() {
        const jobNoInput = document.querySelector('input[name="job_no"], input[name="job_number"], input[id="job_no"]');
        if (!jobNoInput || !jobNoInput.value.trim()) {
            return;
        }
        const jobNo = jobNoInput.value.trim();
        
        try {
            const res = await fetch(`/api/fetch-job-data?job_no=${encodeURIComponent(jobNo)}`);
            const json = await res.json();
            
            if (json.status === 'success') {
                if (json.source === 'excel') {
                    const data = json.data;
                    const labels = document.querySelectorAll('label');
                    labels.forEach(label => {
                        const labelText = label.innerText.trim().toUpperCase();
                        const cleanLabel = labelText.replace(/[^A-Z0-9 ]/g, '').trim();
                        for (const [excelKey, excelVal] of Object.entries(data)) {
                            const cleanKey = excelKey.replace(/[^A-Z0-9 ]/g, '').trim();
                            if (cleanLabel === cleanKey || cleanLabel.includes(cleanKey) || cleanKey.includes(cleanLabel)) {
                                const parent = label.closest('.input-group') || label.parentElement;
                                if (parent) {
                                    const input = parent.querySelector('input[type="text"], input[type="number"], select');
                                    if (input && !input.value) {
                                        input.value = excelVal;
                                        input.dispatchEvent(new Event('input', { bubbles: true }));
                                        input.dispatchEvent(new Event('change', { bubbles: true }));
                                    }
                                }
                            }
                        }
                    });
                }
            }
        } catch (err) {
            console.error("Auto-fetch excel error:", err);
        }
    }

    document.addEventListener('DOMContentLoaded', () => {
        const jobNoInput = document.querySelector('input[name="job_no"], input[name="job_number"], input[id="job_no"]');
        if (jobNoInput) {
            let lastFetched_fetchExcelData = '';
            const autoTrigger_fetchExcelData = () => {
                const v = jobNoInput.value.trim();
                if (v && v !== lastFetched_fetchExcelData) {
                    lastFetched_fetchExcelData = v;
                    fetchExcelData();
                }
            };
            jobNoInput.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    autoTrigger_fetchExcelData();
                }
            });
            jobNoInput.addEventListener('blur', autoTrigger_fetchExcelData);
            jobNoInput.addEventListener('change', autoTrigger_fetchExcelData);
        }
    });
</script>"""

if clean_forwarder in text:
    text = text.replace(clean_forwarder, og_fetch_excel)
    print('SUCCESS: Restored fetchExcelData block')

# 4. Restore window.fetchJobNoData to async function fetchJobNoData()
text = text.replace('window.fetchJobNoData = async function fetchJobNoData() {', 'async function fetchJobNoData() {')

with open(di_path, 'w', encoding='utf-8') as f:
    f.write(text)

print('SUCCESS: Restored digital_indicator.html to OG state!')
