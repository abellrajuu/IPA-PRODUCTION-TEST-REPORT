import sys, os

dash_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\review_dashboard.html'
with open(dash_path, 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Add CSS for btn-docx and category chips
css_target = ".btn-view:hover { background: #0369a1; transform: translateY(-1px); }"
css_replacement = """.btn-view:hover { background: #0369a1; transform: translateY(-1px); }
        .btn-docx { display: inline-flex; align-items: center; gap: 6px; background: #f0f9ff; color: #0284c7; border: 1.5px solid #bae6fd; padding: 6px 12px; border-radius: 8px; font-size: 12.5px; font-weight: 700; text-decoration: none; transition: all 0.2s ease; cursor: pointer; }
        .btn-docx:hover { background: #0284c7; color: #ffffff; border-color: #0284c7; box-shadow: 0 4px 12px rgba(2, 132, 199, 0.25); transform: translateY(-1px); }
        .cat-filter-chips { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }
        .cat-chip { padding: 5px 12px; border-radius: 999px; border: 1px solid #e2e8f0; background: #ffffff; color: #64748b; font-size: 12px; font-weight: 700; cursor: pointer; transition: all 0.15s ease; }
        .cat-chip:hover { border-color: #0284c7; color: #0284c7; background: #f0f9ff; }
        .cat-chip.active { background: #0284c7; color: #ffffff; border-color: #0284c7; box-shadow: 0 2px 8px rgba(2, 132, 199, 0.25); }"""

if css_target in text:
    text = text.replace(css_target, css_replacement)

# 2. Add category filter chips markup right after filter-tabs
filter_markup_target = """<button type="button" id="tabAllJobs" class="filter-tab" onclick="switchFilter('all')">
                <i class="fa-solid fa-layer-group"></i> All Saved Reports
            </button>
        </div>"""

filter_markup_replacement = """<button type="button" id="tabAllJobs" class="filter-tab" onclick="switchFilter('all')">
                <i class="fa-solid fa-layer-group"></i> All Saved Reports
            </button>
        </div>

        <!-- CATEGORY QUICK FILTER CHIPS -->
        <div class="cat-filter-chips" id="catFilterChips">
            <button type="button" class="cat-chip active" onclick="switchCatFilter('all')">All Categories</button>
            <button type="button" class="cat-chip" onclick="switchCatFilter('digital_indicator')"><i class="fa-solid fa-microchip"></i> Digital Indicator</button>
            <button type="button" class="cat-chip" onclick="switchCatFilter('belt_scale')"><i class="fa-solid fa-truck-ramp-box"></i> Belt Scale</button>
            <button type="button" class="cat-chip" onclick="switchCatFilter('crane_scale')"><i class="fa-solid fa-weight-hanging"></i> Crane Scale</button>
            <button type="button" class="cat-chip" onclick="switchCatFilter('weigh_feeder')"><i class="fa-solid fa-boxes-stacked"></i> Weigh Feeder</button>
            <button type="button" class="cat-chip" onclick="switchCatFilter('signal_conditioner')"><i class="fa-solid fa-wave-square"></i> Signal Conditioner</button>
            <button type="button" class="cat-chip" onclick="switchCatFilter('other')"><i class="fa-solid fa-ellipsis"></i> Other Equipment</button>
        </div>"""

if filter_markup_target in text:
    text = text.replace(filter_markup_target, filter_markup_replacement)

# 3. Update JS for humanTypeMap, catFilter, and dual download buttons
js_old = """let allRecords = [];
        let activeTab = 'my_jobs';
        const userRole = "{{ user.role }}";
        const currentUserFullName = "{{ user.full_name }}";
        const currentUsername = "{{ user.username }}";
        const canApprove = ['Verifying Engineer', 'HOD', 'Admin'].includes(userRole);"""

js_new = """let allRecords = [];
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

if js_old in text:
    text = text.replace(js_old, js_new)

# 4. Update applyFilters to also filter by activeCat
apply_old = """if (query) {
                filtered = filtered.filter(r => {"""

apply_new = """if (activeCat !== 'all') {
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

            if (query) {
                filtered = filtered.filter(r => {"""

if apply_old in text:
    text = text.replace(apply_old, apply_new)

# 5. Update renderRecords to output humanType and dual download buttons
render_old = """                    <td><span style="background: #f1f5f9; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 700; color: #0284c7; border: 1px solid #e2e8f0;">${rType}</span></td>
                    <td style="font-weight: 600;">${testedByDisp}</td>
                    <td><span class="badge ${isApproved ? 'badge-approved' : 'badge-pending'}"><i class="fa-solid ${isApproved ? 'fa-check-circle' : 'fa-clock'}"></i> ${statusDisp}</span></td>
                    <td style="text-align: right;">
                        <a href="/download/${encodeURIComponent(r.docx_filename.replace('.docx', '.pdf'))}?v=${Date.now()}" target="_blank" class="btn-view"><i class="fa-solid fa-file-pdf"></i> View PDF</a>
                        ${approveBtn}
                    </td>"""

render_new = """                    <td><span style="background: #f0f9ff; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 700; color: #0284c7; border: 1px solid #bae6fd; display: inline-flex; align-items: center; gap: 5px;"><i class="fa-solid fa-microchip" style="font-size:10px;"></i> ${HUMAN_TYPE_MAP[rType] || rType.replace(/_/g, ' ')}</span></td>
                    <td style="font-weight: 600;">${testedByDisp}</td>
                    <td><span class="badge ${isApproved ? 'badge-approved' : 'badge-pending'}"><i class="fa-solid ${isApproved ? 'fa-check-circle' : 'fa-clock'}"></i> ${statusDisp}</span></td>
                    <td style="text-align: right;">
                        <div style="display: flex; align-items: center; justify-content: flex-end; gap: 8px;">
                            <a href="/download/${encodeURIComponent(r.docx_filename.replace('.docx', '.pdf'))}?v=${Date.now()}" target="_blank" class="btn-view" title="View / Print PDF"><i class="fa-solid fa-file-pdf"></i> PDF</a>
                            <a href="/download/${encodeURIComponent(r.docx_filename)}" class="btn-docx" title="Download Master Word Document"><i class="fa-solid fa-file-word"></i> Word</a>
                            ${approveBtn}
                        </div>
                    </td>"""

if render_old in text:
    text = text.replace(render_old, render_new)

with open(dash_path, 'w', encoding='utf-8') as f:
    f.write(text)

print('SUCCESS: Updated review_dashboard.html with UI-UX Pro Max Data-Dense Dashboard!')
