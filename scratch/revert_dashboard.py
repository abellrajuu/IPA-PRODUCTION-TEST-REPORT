import sys, os

dash_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\review_dashboard.html'
with open(dash_path, 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Remove added CSS
css_added = """.btn-docx { padding: 8px 14px; border-radius: 8px; background: #f8fafc; color: #1e40af; text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; font-size: 13px; border: 1.5px solid #cbd5e1; transition: all 0.2s; cursor: pointer; }
        .btn-docx:hover { background: #1e40af; color: #fff; border-color: #1e40af; box-shadow: 0 4px 12px rgba(30, 64, 175, 0.25); transform: translateY(-1px); }
        .cat-filter-chips { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin: 16px 0 20px; }
        .cat-chip { padding: 7px 16px; border-radius: 999px; border: 1.5px solid #e2e8f0; background: #ffffff; color: #64748b; font-size: 12.5px; font-weight: 700; cursor: pointer; transition: all 0.2s ease; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04); }
        .cat-chip:hover { border-color: #0284c7; color: #0284c7; background: #f0f9ff; transform: translateY(-1px); }
        .cat-chip.active { background: #0284c7; color: #ffffff; border-color: #0284c7; box-shadow: 0 3px 10px rgba(2, 132, 199, 0.25); }"""

if css_added in text:
    text = text.replace(css_added, '')
    print('SUCCESS: Removed added CSS from review_dashboard.html')

# 2. Remove category quick filter chips markup
chips_start = text.find('<!-- CATEGORY QUICK FILTER CHIPS -->')
if chips_start != -1:
    chips_end = text.find('</div>', chips_start) + 6
    text = text[:chips_start] + text[chips_end:]
    print('SUCCESS: Removed cat filter chips markup')

# 3. Restore JS
js_og = """let allRecords = [];
        let activeTab = 'my_jobs';
        const userRole = "{{ user.role }}";
        const currentUserFullName = "{{ user.full_name }}";
        const currentUsername = "{{ user.username }}";
        const canApprove = ['Verifying Engineer', 'HOD', 'Admin'].includes(userRole);"""

js_modified = """let allRecords = [];
        let activeTab = 'my_jobs';
        let activeCat = 'all';
        const userRole = "{{ user.role }}";
        const currentUserFullName = "{{ user.full_name }}";
        const currentUsername = "{{ user.username }}";
        const canApprove = ['Verifying Engineer', 'HOD', 'Admin'].includes(userRole);

        const HUMAN_TYPE_MAP = {
            'digital_indicator': 'Digital Indicator',
            'belt_scale': 'Belt Scale',
            'crane_scale': 'Crane Scale',
            'weigh_feeder': 'Weigh Feeder',
            'loss_in_weigh_feeder': 'Loss in Weigh Feeder',
            'signal_conditioner': 'Signal Conditioner',
            'batching_system': 'Batching System',
            'remote_indicator': 'Remote Indicator',
            'misc_report': 'Misc Report',
            'job_traveller_card': 'Job Traveller Card',
            'odd_system': 'ODD System',
            'dd_system': 'DD System',
            'trip_safe': 'Trip Safe',
            'vibration_switch': 'Vibration Switch',
            'vibration_meter': 'Vibration Meter',
            'acc_mv': 'ACC mV',
            'acc_charge': 'ACC Charge',
            'charge_amplifier': 'Charge Amplifier',
            'inprocess': 'Inprocess Register'
        };

        function switchCatFilter(cat) {
            activeCat = cat;
            document.querySelectorAll('.cat-chip').forEach(btn => {
                const isMatch = btn.getAttribute('onclick').includes(`'${cat}'`);
                btn.classList.toggle('active', isMatch);
            });
            applyFilters();
        }"""

if js_modified in text:
    text = text.replace(js_modified, js_og)
    print('SUCCESS: Restored original JS in review_dashboard.html')

# 4. Restore applyFilters
apply_modified = """if (activeCat !== 'all') {
                filtered = filtered.filter(r => {
                    let rType = '';
                    try {
                        let jd = JSON.parse(r.data_json);
                        rType = (jd.report_type || jd.system_name || '').toLowerCase();
                    } catch(e){}
                    if (activeCat === 'other') {
                        return !['digital_indicator', 'belt_scale', 'crane_scale', 'weigh_feeder', 'signal_conditioner'].includes(rType);
                    }
                    return rType.includes(activeCat);
                });
            }

            if (query) {"""

apply_og = """if (query) {"""

if apply_modified in text:
    text = text.replace(apply_modified, apply_og)
    print('SUCCESS: Restored original applyFilters')

# 5. Restore renderRecords
render_modified = """                    <td><span style="background: #f0f9ff; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 700; color: #0284c7; border: 1px solid #bae6fd; display: inline-flex; align-items: center; gap: 5px;"><i class="fa-solid fa-microchip" style="font-size:10px;"></i> ${HUMAN_TYPE_MAP[rType] || rType.replace(/_/g, ' ')}</span></td>
                    <td style="font-weight: 600;">${testedByDisp}</td>
                    <td><span class="badge ${isApproved ? 'badge-approved' : 'badge-pending'}"><i class="fa-solid ${isApproved ? 'fa-check-circle' : 'fa-clock'}"></i> ${statusDisp}</span></td>
                    <td style="text-align: right;">
                        <div style="display: flex; align-items: center; justify-content: flex-end; gap: 8px;">
                            <a href="/download/${encodeURIComponent(r.docx_filename.replace('.docx', '.pdf'))}?v=${Date.now()}" target="_blank" class="btn-view" title="View / Print PDF"><i class="fa-solid fa-file-pdf"></i> PDF</a>
                            <a href="/download/${encodeURIComponent(r.docx_filename)}" class="btn-docx" title="Download Master Word Document"><i class="fa-solid fa-file-word"></i> Word</a>
                            ${approveBtn}
                        </div>
                    </td>"""

render_og = """                    <td><span style="background: #f1f5f9; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 700; color: #0284c7; border: 1px solid #e2e8f0;">${rType}</span></td>
                    <td style="font-weight: 600;">${testedByDisp}</td>
                    <td><span class="badge ${isApproved ? 'badge-approved' : 'badge-pending'}"><i class="fa-solid ${isApproved ? 'fa-check-circle' : 'fa-clock'}"></i> ${statusDisp}</span></td>
                    <td style="text-align: right;">
                        <a href="/download/${encodeURIComponent(r.docx_filename.replace('.docx', '.pdf'))}?v=${Date.now()}" target="_blank" class="btn-view"><i class="fa-solid fa-file-pdf"></i> View PDF</a>
                        ${approveBtn}
                    </td>"""

if render_modified in text:
    text = text.replace(render_modified, render_og)
    print('SUCCESS: Restored original renderRecords')

with open(dash_path, 'w', encoding='utf-8') as f:
    f.write(text)

print('SUCCESS: Restored review_dashboard.html to OG state!')
