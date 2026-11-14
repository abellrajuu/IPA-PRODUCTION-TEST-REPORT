// Deprecated wrapper removed. The submit handler will be assigned after its definition.



function escapeHtml(str) {
    if (str === undefined || str === null) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
window.escapeHtml = escapeHtml;

// Auto-highlight key technical specifications & hardware models (Load Cell Model, Indicator Model, Tag No, etc.)
function applyKeySpecHighlights() {
    const keyPatterns = [
        /load\s*cell\s*model/i,
        /sensor.*model/i,
        /digital\s*indicator\s*model/i,
        /indicator\s*model/i,
        /di[_\s]*model/i,
        /lc[_\s]*model/i,
        /conveyor\s*no/i,
        /tag\s*no/i,
        /crane\s*id/i,
        /tracking\s*no/i,
        /scale\s*no/i,
        /capacity/i,
        /job\s*no/i,
        /customer/i,
        /junction\s*box\s*model/i,
        /jbox\s*model/i,
        /firmware\s*sw\s*version/i
    ];
    const keyFieldNames = ['lc_model', 'di_model', 'capacity', 'job_no', 'customer', 'conveyor_no', 'crane_id', 'tracking_no', 'jbox_model', 'sw_version'];

    document.querySelectorAll('.input-group:not([data-key-highlight]), .form-group:not([data-key-highlight])').forEach(group => {
        let isMatch = false;
        const label = group.querySelector('label');
        if (label) {
            const txt = label.textContent.trim();
            isMatch = keyPatterns.some(rx => rx.test(txt));
        }
        if (!isMatch) {
            const input = group.querySelector('input, select');
            if (input) {
                const name = (input.name || '').toLowerCase();
                const id = (input.id || '').toLowerCase();
                if (keyFieldNames.includes(name) || keyFieldNames.includes(id)) {
                    isMatch = true;
                }
            }
        }
        if (isMatch) {
            group.setAttribute('data-key-highlight', 'true');
            group.classList.add('highlight-spec');
        } else {
            group.setAttribute('data-key-highlight', 'false');
        }
    });
}
window.applyKeySpecHighlights = applyKeySpecHighlights;

document.addEventListener('DOMContentLoaded', () => {
    applyKeySpecHighlights();
    
    window.closeDownloadBanner = function() {
        const banner = document.getElementById("downloadBanner");
        if (banner) {
            banner.classList.remove("show-banner");
            banner.style.setProperty("display", "none", "important");
        }
    };

    // Ensure download banner is strictly hidden on page load
    const initialBanner = document.getElementById("downloadBanner");
    if (initialBanner) {
        initialBanner.classList.remove("show-banner");
        initialBanner.style.setProperty("display", "none", "important");
    }

    const CURRENT_BUILD_VERSION = document.querySelector('meta[name="app-build-version"]')?.getAttribute('content') || '';

    // STEP-BY-STEP WIZARD STATE
    const hasWizard = document.querySelectorAll('.step-panel').length > 0 || document.getElementById('step-box-1') !== null;
    const stepPanels = document.querySelectorAll('.step-panel').length > 0 ? document.querySelectorAll('.step-panel') : (hasWizard ? document.querySelectorAll('.section-box') : []);
    const totalSteps = stepPanels.length > 0 ? stepPanels.length : 1;

    let currentStep = 1;

    // Elements
    const btnNextStep = document.getElementById('btnNextStep');
    const btnPrevStep = document.getElementById('btnPrevStep');
    const btnSubmitPDF = document.getElementById('btnSubmitPDF');
    const progressFill = document.getElementById('wizardProgressFill');
    const stepIndicators = document.querySelectorAll('.step-indicator');
    
    const btnTopAutoFill = document.getElementById('btnTopAutoFill');
    const numReadingsSelect = document.getElementById('num_readings_select');

    // OFFLINE STORAGE & WI-FI SERVER CONFIG MANAGEMENT
    const OFFLINE_STORAGE_KEY = 'belt_scale_offline_records_v1';
    let isServerOnline = false;

    async function downloadWithSaveAsPrompt(fileUrl, suggestedFilename) {
        if (!fileUrl) return;
        const fname = suggestedFilename || (fileUrl.toLowerCase().endsWith('.docx') ? "Report.docx" : "Report.pdf");
        let fullUrl = fileUrl;
        if (!fullUrl.startsWith("http://") && !fullUrl.startsWith("https://")) {
            const base = (typeof getServerBaseUrl === 'function' ? getServerBaseUrl() : window.location.origin).replace(/\/+$/, '');
            fullUrl = `${base}${fullUrl.startsWith('/') ? '' : '/'}${fullUrl}`;
        }

        const isDocx = fname.toLowerCase().endsWith('.docx');
        const mimeType = isDocx 
            ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
            : 'application/pdf';

        // 1. Android Native App DownloadManager Interface (direct device download)
        if (window.AndroidNative && typeof window.AndroidNative.downloadFile === 'function') {
            try {
                window.AndroidNative.downloadFile(fullUrl, fname, mimeType);
                return;
            } catch (err) {
                console.warn("AndroidNative downloadFile failed, fallback:", err);
            }
        }

        const isAndroidApp = /NativeApp|ProductionReportApp/i.test(navigator.userAgent);
        if (isAndroidApp) {
            // Trigger Android WebView's DownloadListener via hidden iframe
            const ifr = document.createElement('iframe');
            ifr.style.display = 'none';
            ifr.src = fullUrl;
            document.body.appendChild(ifr);
            setTimeout(() => { try { document.body.removeChild(ifr); } catch(e){} }, 4000);
            return;
        }

        // 2. Fetch as Blob and trigger automatic instant download (standard browser download directly to device)
        try {
            const resp = await fetch(fullUrl, { credentials: 'include' });
            if (resp.ok) {
                const blob = await resp.blob();
                const blobUrl = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.style.display = "none";
                a.href = blobUrl;
                a.download = fname;
                document.body.appendChild(a);
                a.click();
                setTimeout(() => {
                    window.URL.revokeObjectURL(blobUrl);
                    try { document.body.removeChild(a); } catch(e){}
                }, 5000);
                return;
            }
        } catch (err) {
            console.warn("Blob fetch download failed, fallback to direct anchor:", err);
        }

        // 3. Fallback direct anchor download
        try {
            const a = document.createElement("a");
            a.style.display = "none";
            a.href = fullUrl;
            a.download = fname;
            a.target = "_blank";
            document.body.appendChild(a);
            a.click();
            setTimeout(() => { try { document.body.removeChild(a); } catch(e){} }, 2000);
        } catch(e) {
            console.error("Direct anchor download error:", e);
        }
    }
    window.downloadWithSaveAsPrompt = downloadWithSaveAsPrompt;

    function getServerBaseUrl() {
        if (window.location.protocol.startsWith('http')) {
            return window.location.origin;
        }
        const savedServerIp = localStorage.getItem('production_server_ip') || '127.0.0.1:5005';
        return savedServerIp.startsWith('http') ? savedServerIp : `http://${savedServerIp}`;
    }

    function getOfflineRecords() {
        try {
            return JSON.parse(localStorage.getItem(OFFLINE_STORAGE_KEY) || '[]');
        } catch(e) { return []; }
    }

    function saveOfflineRecord(payload) {
        const records = getOfflineRecords();
        payload._offline_id = 'OFFLINE_' + Date.now();
        payload._created_at = new Date().toLocaleString();
        records.push(payload);
        localStorage.setItem(OFFLINE_STORAGE_KEY, JSON.stringify(records));
        updateOfflineBadge();
        showToast('Saved Locally on Device (Offline Mode)', 'warning');
    }

    function removeOfflineRecord(offlineId) {
        let records = getOfflineRecords();
        records = records.filter(r => r._offline_id !== offlineId);
        localStorage.setItem(OFFLINE_STORAGE_KEY, JSON.stringify(records));
        updateOfflineBadge();
    }

    function updateOfflineBadge() {
        const count = getOfflineRecords().length;
        const badge = document.getElementById('offlineQueueBadge');
        if (badge) {
            if (count > 0) {
                badge.style.display = 'inline-flex';
                badge.style.cursor = 'pointer';
                badge.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i> ${count} Pending Sync (Tap to Sync Now)`;
                badge.onclick = () => {
                    showToast('Syncing offline records to server...', 'warning');
                    syncOfflineRecords();
                };
            } else {
                badge.style.display = 'none';
            }
        }
    }

    // AUTOMATIC SYNC WHEN BACK IN SERVER RANGE
    async function syncOfflineRecords() {
        const records = getOfflineRecords();
        if (records.length === 0) {
            updateOfflineBadge();
            return;
        }

        let syncedCount = 0;
        for (const record of records) {
            try {
                const res = await fetch(`${getServerBaseUrl()}/api/submit`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(record)
                });
                const data = await res.json();
                if (data.status === 'success') {
                    removeOfflineRecord(record._offline_id);
                    syncedCount++;
                }
            } catch(e) {
                break; // Server went offline again
            }
        }

        if (syncedCount > 0) {
            showToast(`Auto-Synced ${syncedCount} Offline Records to Server Database`, 'success');
            loadHistory();
        }
        updateOfflineBadge();
    }

    // SERVER ONLINE HEARTBEAT MONITOR
    async function checkConnectivity() {
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 2500);
            const res = await fetch(`${getServerBaseUrl()}/api/version`, { signal: controller.signal });
            clearTimeout(timeoutId);
            if (res.ok) {
                const wasOffline = !isServerOnline;
                isServerOnline = true;
                if (wasOffline) {
                    showToast('Connected to Server Network', 'success');
                }
                res.clone().json().then(verData => {
                    if (verData && verData.version && !window.__updateNotified) {
                        const localBuild = document.querySelector('meta[name="app-build-version"]')?.content;
                        if (localBuild && verData.build_version && localBuild !== verData.build_version) {
                            window.__updateNotified = true;
                            if (typeof showToast === 'function') {
                                showToast(`🚀 App update available (v${verData.version}). Tap refresh to load latest version.`, 'info');
                            }
                        }
                    }
                }).catch(() => {});
                if (getOfflineRecords().length > 0) {
                    syncOfflineRecords();
                }
            } else {
                isServerOnline = false;
            }
        } catch(e) {
            isServerOnline = false;
        }
        updateNetworkStatusIndicator();
        updateOfflineBadge();
    }

    function updateNetworkStatusIndicator() {
        const indicator = document.getElementById('networkStatusIndicator');
        if (indicator) {
            if (isServerOnline) {
                indicator.className = 'user-badge';
                indicator.style.borderColor = 'rgba(16, 185, 129, 0.4)';
                indicator.style.color = '#34d399';
                indicator.innerHTML = '<span class="pulse-dot"></span> Server Connected';
            } else {
                indicator.className = 'user-badge';
                indicator.style.borderColor = 'rgba(245, 158, 11, 0.5)';
                indicator.style.color = '#fbbf24';
                indicator.style.background = 'rgba(245, 158, 11, 0.15)';
                indicator.innerHTML = '<i class="fa-solid fa-wifi-slash"></i> Offline Mode (Device Storage)';
            }
        }
    }

    setInterval(checkConnectivity, 30000);
    window.addEventListener('online', checkConnectivity);
    window.addEventListener('offline', updateNetworkStatusIndicator);
    checkConnectivity();

    window.goToStep = function(stepNum) {
        const panels = document.querySelectorAll('.step-panel');
        let detectedSteps = panels.length;
        const hasStepBoxes = document.getElementById('step-box-1') || document.getElementById('step-box-2');
        if (hasStepBoxes) {
            let boxCount = 0;
            for (let i = 1; i <= 10; i++) {
                if (document.getElementById('step-box-' + i) || document.getElementById('step-' + i)) {
                    boxCount++;
                }
            }
            if (boxCount > 0) detectedSteps = boxCount;
        }
        const maxSteps = Math.max(1, detectedSteps || totalSteps);
        const clampedStep = Math.max(1, Math.min(maxSteps, parseInt(stepNum, 10) || 1));
        currentStep = clampedStep;

        if (hasStepBoxes) {
            for (let i = 1; i <= 10; i++) {
                const box = document.getElementById('step-box-' + i) || document.getElementById('step-' + i);
                const st = document.getElementById('st-' + i) || document.querySelector(`.step-indicator[data-step="${i}"]`);
                if (box) {
                    if (i === clampedStep) {
                        box.classList.add('active-step');
                        box.classList.add('active');
                        box.style.display = 'block';
                    } else {
                        box.classList.remove('active-step');
                        box.classList.remove('active');
                        box.style.display = 'none';
                    }
                }
                if (st) {
                    if (i === clampedStep) {
                        st.classList.add('active');
                        st.classList.remove('completed');
                    } else if (i < clampedStep) {
                        st.classList.remove('active');
                        st.classList.add('completed');
                    } else {
                        st.classList.remove('active');
                        st.classList.remove('completed');
                    }
                }
            }
        } else {
            panels.forEach((p, idx) => {
                if (idx + 1 === clampedStep) {
                    p.classList.add('active');
                    p.style.display = 'block';
                } else {
                    p.classList.remove('active');
                    p.style.display = 'none';
                }
            });
            stepIndicators.forEach((st, idx) => {
                const stNum = parseInt(st.getAttribute('data-step'), 10) || (idx + 1);
                if (stNum === clampedStep) {
                    st.classList.add('active');
                    st.classList.remove('completed');
                } else if (stNum < clampedStep) {
                    st.classList.remove('active');
                    st.classList.add('completed');
                } else {
                    st.classList.remove('active');
                    st.classList.remove('completed');
                }
            });
        }

        if (progressFill) {
            progressFill.style.width = `${Math.round((clampedStep / maxSteps) * 100)}%`;
        }
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };



    

    // INITIALIZE WIZARD ON PAGE LOAD
    if (hasWizard && stepPanels.length > 0) {
        goToStep(1);
    }

    if (btnNextStep) {
        btnNextStep.addEventListener('click', () => {
            goToStep(currentStep + 1);
        });
    }

    if (btnPrevStep) {
        btnPrevStep.addEventListener('click', () => {
            goToStep(currentStep - 1);
        });
    }

    stepIndicators.forEach(ind => {
        ind.addEventListener('click', () => {
            const targetStep = parseInt(ind.getAttribute('data-step'));
            goToStep(targetStep);
        });
    });

    // Routine Tests Definition
    const defaultRoutineTests = [
        { num: '1.', name: 'Input Supply Voltage (+/-10%)', spec: '230 V AC', act: '229.4 V AC' },
        { num: '2.', name: 'Derived DC Voltages (+/-5%)', spec: '+24 V DC', act: '+24.1 V DC' },
        { num: '3.', name: 'Display and Keypad Functionality', spec: 'OK', act: 'OK Verified' },
        { num: '4.', name: 'Belt Load as specified in G.A. or MPL', spec: '150.0 Kg/m', act: '149.6 Kg/m' },
        { num: '5.', name: 'Belt Speed as specified in G.A. or MPL', spec: '2.0 m/s', act: '2.0 m/s' },
        { num: '6.', name: 'Number of PF Contacts', spec: '4 Contacts', act: '4 Contacts Passed' },
        { num: '7.', name: 'Communication Output (Specify)', spec: 'RS485 Modbus', act: 'Connected 9600-8-N-1' },
        { num: '8.', name: 'Analog Output (Specify number)', spec: '4-20 mA', act: '4.02 mA to 19.98 mA' },
        { num: '9.', name: 'Wiring, TB and Component Layout', spec: 'OK Drawing', act: 'OK Checked' }
    ];

    // Initialize Routine Tests Table
    const routineTableBody = document.getElementById('routineTestsBody');
    if (routineTableBody) {
        routineTableBody.innerHTML = defaultRoutineTests.map((t, idx) => `
            <tr>
                <td><strong>${t.num}</strong></td>
                <td>${t.name}</td>
                <td><input type="text" id="t_spec_${idx}" value="${t.spec}" placeholder="Specified"></td>
                <td><input type="text" id="t_act_${idx}" value="${t.act}" placeholder="Actual"></td>
            </tr>
        `).join('');
    }

    // Initialize 11-Row Measurement Grid Table
    const gridTableBody = document.getElementById('gridMatrixBody');
    if (gridTableBody) {
        let rowsHtml = '';
        for (let i = 0; i < 11; i++) {
            rowsHtml += `
                <tr id="grid_row_${i}">
                    <td><input type="text" id="g_${i}_0" placeholder="Load kg"></td>
                    <td><input type="text" id="g_${i}_1" placeholder="Belt Kg/m"></td>
                    <td><input type="text" id="g_${i}_2" placeholder="Speed m/s"></td>
                    <td><input type="text" id="g_${i}_3" placeholder="Rate tph"></td>
                    <td><input type="text" id="g_${i}_4" placeholder="Totalizer Tonnes"></td>
                    <td><input type="text" id="g_${i}_5" placeholder="O/p mA"></td>
                </tr>
            `;
        }
        gridTableBody.innerHTML = rowsHtml;
    }

    // Dynamic Readings Row Visibility Handler - ensures all rows (including 6-10) are always editable
    function updateReadingsVisibility() {
        for (let i = 0; i < 11; i++) {
            const rowElem = document.getElementById(`grid_row_${i}`);
            if (!rowElem) continue;

            const inputs = rowElem.querySelectorAll('input');
            rowElem.style.opacity = '1';
            rowElem.style.pointerEvents = 'auto';
            inputs.forEach(inp => {
                inp.disabled = false;
            });
        }
    }

    if (numReadingsSelect) {
        numReadingsSelect.addEventListener('change', updateReadingsVisibility);
    }
    updateReadingsVisibility();

    // Set Default Date to Today
    const dateInput = document.getElementById('date');
    if (dateInput) {
        dateInput.value = new Date().toISOString().split('T')[0];
    }

    // Auto calculate RPM when Speed (S) and Wheel Dia (D) change
    const speedInput = document.getElementById('tacho_belt_speed');
    const diaInput = document.getElementById('tacho_wheel_dia');
    const rpmInput = document.getElementById('tacho_rpm');

    function calcRpm() {
        if (!speedInput || !diaInput || !rpmInput) return;
        const S = parseFloat(speedInput.value);
        const D = parseFloat(diaInput.value);
        if (!isNaN(S) && !isNaN(D) && D > 0) {
            const rpm = (S / (3.14159 * D)) * 60;
            rpmInput.value = rpm.toFixed(1) + ' RPM';
        }
    }
    if (speedInput) speedInput.addEventListener('input', calcRpm);
    if (diaInput) diaInput.addEventListener('input', calcRpm);

    // Capture local template-specific filler if already declared by the active page
    const localTemplateFiller = (typeof window.fillDemoData === 'function')
        ? window.fillDemoData
        : (typeof window.fillSampleTestData === 'function'
            ? window.fillSampleTestData
            : (typeof window.fillSampleData === 'function' ? window.fillSampleData : null));

    // =========================================================================
    // UNIVERSAL AUTO-FILL SYSTEM FOR ALL 15+ INDUSTRIAL TESTING TEMPLATES
    // Populates 100% of fields, steps, checklists, tables & triggers math logic
    // =========================================================================
    function triggerUniversalAutoFill() {
        const setVal = (idOrName, val) => {
            if (val === undefined || val === null) return;
            const el = document.getElementById(idOrName) || document.querySelector(`[name="${idOrName}"]`);
            if (el) {
                el.value = val;
                try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch(e){}
                try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
            }
        };

        // 1. If active template defined its own dedicated demo filler, invoke it dynamically
        const pageCustomFiller = (typeof window.fillDemoData === 'function' && window.fillDemoData !== triggerUniversalAutoFill)
            ? window.fillDemoData
            : (typeof window.fillSampleTestData === 'function' && window.fillSampleTestData !== triggerUniversalAutoFill)
                ? window.fillSampleTestData
                : (typeof window.fillSampleData === 'function' && window.fillSampleData !== triggerUniversalAutoFill)
                    ? window.fillSampleData
                    : (typeof localTemplateFiller === 'function' && localTemplateFiller !== triggerUniversalAutoFill ? localTemplateFiller : null);

        if (typeof pageCustomFiller === 'function') {
            try {
                pageCustomFiller();
                if (typeof showToast === 'function') {
                    showToast('⚡ Auto-Fill Completed! All parameters populated.', 'success');
                }
                return;
            } catch(e) {
                console.warn("Local template filler warning, falling back to universal sweep:", e);
            }
        }

        // Determine active report type / template
        let reportType = (document.querySelector('input[name="report_type"]')?.value || '').toLowerCase().trim();
        if (!reportType) {
            const path = window.location.pathname.toLowerCase().replace(/-/g, '_');
            for (const key of ['belt_scale', 'weigh_feeder', 'crane_scale', 'digital_indicator', 'batching_system',
                               'loss_in_weigh_feeder', 'signal_conditioner', 'trip_safe', 'vibration_switch',
                               'vibration_meter', 'acc_charge', 'acc_mv', 'charge_amplifier', 'remote_indicator',
                               'dd_system', 'odd_system', 'job_traveller_card', 'misc_report', 'equipment_list',
                               'work_instructions', 'inprocess_register', 'inprocess', 'performance_delay', 'performance_index']) {
                if (path.includes(key)) { reportType = key; break; }
            }
        }

        const randJob = Math.floor(100 + Math.random() * 900);
        const todayIso = new Date().toISOString().split('T')[0];
        const clientList = ['UltraTech Cement Ltd', 'Tata Steel Plant', 'JSW Steel Mining', 'Adani Power Ltd', 'Jindal Steel & Power Ltd'];
        const chosenClient = clientList[Math.floor(Math.random() * clientList.length)];

        // 2. Targeted System-Specific Precision Filling
        if (reportType === 'weigh_feeder') {
            setVal('job_no', `JOB-WF-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('capacity', '100 TPH');
            setVal('material', 'Limestone / Clinker');
            setVal('res_load', '0.01 kg/m');
            setVal('res_speed', '0.001 m/s');
            setVal('res_rate', '0.01 tph');
            setVal('res_tot', '0.001 t');
            setVal('wf_model', 'WF-800MM-HD');
            setVal('wf_serial', `SN-WF-88${randJob % 10}`);
            setVal('lp_model', 'LP-PBS-100');
            setVal('lp_serial', `SN-LP-90${randJob % 10}`);
            setVal('ctrl_model', 'WF-CTRL-PRO');
            setVal('ctrl_serial', `SN-CTRL-77${randJob % 10} (v3.2)`);
            setVal('lc_model', 'LC-BEAM-350');
            setVal('lc_serial_1', `SN-LC-01, SN-LC-02`);
            setVal('tacho_model', 'TACHO-PULSE-1000');
            setVal('jbox_1_model', 'JBOX-4IN1');
            setVal('drive_make', 'ABB');
            setVal('drive_model', 'ACS380-3KW');
            setVal('drive_serial', `SN-DRV-44${randJob % 10}`);
            setVal('drive_kw', '3.7 kW');
            setVal('motor_make', 'Siemens');
            setVal('motor_model', '1LE1001-3.7KW');
            setVal('motor_serial', `SN-MTR-88${randJob % 10}`);
            setVal('motor_kw', '3.7 kW');
            setVal('gb_make', 'SEW Eurodrive');
            setVal('gb_model', 'KA47-3.7KW');
            setVal('gb_serial', `SN-GB-22${randJob % 10}`);
            setVal('ral_make', 'RAL Industrial');
            setVal('ral_model', 'RAL-PRO');
            setVal('ral_serial', `SN-RAL-10${randJob % 10}`);
            setVal('belt_len_width', '12000 / 800 mm');
            setVal('act_speed_measured', '0.40 m/s');
            setVal('spec_1', '230V AC +/-10%'); setVal('act_1', '230V AC');
            setVal('spec_2', '24V DC +/-5%'); setVal('act_2', '24.1V DC');
            setVal('spec_3', 'Display & Keypad OK'); setVal('act_3', 'Passed');
            setVal('spec_4', '15.5 kg/m'); setVal('act_4', '15.5 kg/m');
            setVal('spec_5', '0.40 m/s'); setVal('act_5', '0.40 m/s');
            setVal('calc_speed_10', '0.040 m/s'); setVal('meas_speed_10', '0.040 m/s');
            setVal('calc_speed_25', '0.100 m/s'); setVal('meas_speed_25', '0.100 m/s');
            setVal('calc_speed_50', '0.200 m/s'); setVal('meas_speed_50', '0.200 m/s');
            setVal('calc_speed_75', '0.300 m/s'); setVal('meas_speed_75', '0.300 m/s');
            setVal('calc_speed_100', '0.400 m/s'); setVal('meas_speed_100', '0.400 m/s');
            setVal('cal_chain_wt', '15.5 Kg/m');
            setVal('vol_mode', '10% - 100% Speed Verified OK');
            setVal('grav_mode', '10% - 100% Rate Verified OK');
            setVal('func_mode', 'Local & Remote Mode Verified OK');
            setVal('func_di', 'Digital Inputs Verified OK');
            setVal('func_pf', 'PF Relay Contacts Verified OK');
            setVal('func_aio', '4-20mA I/O Verified OK');
            setVal('func_comm', 'Modbus RTU (Slave: 1, Baud: 9600 8N1) Verified OK');
        } else if (reportType === 'crane_scale') {
            setVal('job_no', `JOB-CWS-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('capacity', '10 Tons');
            setVal('resolution', '1.0 kg');
            setVal('stamp_no', `STAMP-2026-${randJob}`);
            setVal('accuracy', 'Class III');
            setVal('crane_model', 'CWS-5000 Crane Unit');
            setVal('crane_serial', `SN-CWS-99${randJob % 10}`);
            setVal('lc_model', 'LC-CRANE-10T');
            setVal('lc_serial', `SN-LC-88${randJob % 10}`);
            setVal('remote_details', 'Remote RDU-500 Wireless');
            setVal('charger_details', `Charger 12V 2A (SN-CH-${randJob % 100})`);
            setVal('hook_serial', `SN-HK-88${randJob % 10}`);
            setVal('shackle_serial', `SN-SK-77${randJob % 10}`);
            setVal('spec_1', '12V DC Battery / Full Charge'); setVal('act_1', '12.4V DC Passed');
            setVal('spec_2', 'Function of all switches OK'); setVal('act_2', 'Passed');
            setVal('spec_3', 'Wireless Range >= 15m'); setVal('act_3', 'Passed (25m tested)');
            for (let i = 1; i <= 6; i++) {
                const ton = (i - 1) * 2;
                setVal(`cal_load_${i}`, `${ton}.00 T`);
                setVal(`cal_ind_${i}`, `${ton}.00 T`);
                setVal(`cal_err_${i}`, `0.00 %`);
            }
        } else if (reportType === 'digital_indicator') {
            setVal('job_no', `JOB-DI-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('capacity', '100 Tonnes');
            setVal('conveyor_no', `DI-TAG-0${randJob % 9 + 1}`);
            setVal('di_model', 'DI-3000-PRO');
            setVal('di_serial', `SN-DI-88${randJob}`);
            setVal('sw_version', 'v3.2.0');
            setVal('lc_model', 'LC-CZ50T');
            for (let i = 1; i <= 8; i++) setVal(`s${i}`, `LC-SN-00${i}`);
            setVal('jbox_model', 'JB-8P');
            setVal('jbox_serial', `JB-SN-44${randJob % 10}`);
            setVal('short_check', 'OK (>100M Ohm)');
            setVal('spec_1', '230V AC +/-10%'); setVal('act_1', '230V AC');
            setVal('spec_2', '+24V DC, +5V DC'); setVal('act_2', '+24.1V, +5.01V');
            setVal('spec_3', 'Display OK, Keys Responsive'); setVal('act_3', 'Passed');
            setVal('spec_4', 'Full Capacity Span'); setVal('act_4', 'Passed');
            setVal('spec_5', 'Zero Balance OK'); setVal('act_5', 'Passed');
            setVal('check_6', 'Checked: YES'); setVal('check_7', 'Checked: YES');
            setVal('check_8', 'Checked: YES'); setVal('check_9', 'N/A'); setVal('check_10', '2A Fast Blow');
            for (let r = 1; r <= 12; r++) {
                let pct = (r - 1) * 10;
                setVal(`g${r}_0`, `${pct}%`);
                for (let c = 1; c <= 4; c++) setVal(`g${r}_${c}`, `${pct * 10}`);
                setVal(`g${r}_5`, `${(4 + (pct * 16 / 100)).toFixed(2)}`);
                for (let c = 6; c <= 9; c++) setVal(`g${r}_${c}`, `${pct * 10}`);
            }
            for (let lc = 1; lc <= 8; lc++) {
                setVal(`span_lc${lc}`, '10.00');
                setVal(`tare_lc${lc}`, '0.00');
                setVal(`excv_lc${lc}`, '10.0V');
            }
        } else if (reportType === 'batching_system') {
            setVal('job_no', `JOB-BCH-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('capacity', '500 kg');
            setVal('material', 'Fine Aggregate / Sand');
            setVal('res_l', '0.1 kg');
            setVal('bch_model', 'BCH-500-AUTO');
            setVal('bch_serial', `SN-BCH-33${randJob % 10}`);
            setVal('ctrl_model', 'PLC-BATCH-MASTER');
            setVal('ctrl_serial', `SN-PLC-55${randJob % 10}`);
            setVal('sw_version', 'v2.4.1');
            setVal('lc_model', 'LC-S-BEAM-250');
            setVal('s1', 'LC-S1-01'); setVal('s2', 'LC-S2-02');
            setVal('s3', 'LC-S3-03'); setVal('s4', 'LC-S4-04');
            setVal('jbox_model1', 'JB-IP65'); setVal('jbox_serial1', 'JB-SN-101');
            setVal('jbox_model2', 'JB-IP65'); setVal('jbox_serial2', 'JB-SN-102');
            setVal('short_check', 'OK (>100M Ohm)');
            setVal('spec_1', '230V AC +/-10%'); setVal('act_1', '230V AC');
            setVal('spec_2', 'Display function & keyboard OK'); setVal('act_2', 'Passed');
            setVal('spec_3', 'Speed of Bagging / Batching'); setVal('act_3', '120 batches/hr Passed');
            setVal('spec_4', 'Accuracy +/-0.1%'); setVal('act_4', '+/-0.05% Passed');
            setVal('spec_5', 'Bulk Density of Material'); setVal('act_5', '1.6 t/m3');
        } else if (reportType === 'trip_safe') {
            setVal('job_no', `JOB-TS-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('ts_model', 'TS-OVERLOAD-PRO');
            setVal('ts_serial', `SN-TS-44${randJob % 10}`);
            setVal('lc_model', 'LC-COMP-10T');
            setVal('lc_serial', `SN-LC-90${randJob % 10}`);
            const tripLoads = [0, 2000, 4000, 6000, 8000, 10000];
            const leds = ['LED 1 (Green)', 'LED 3 (Green)', 'LED 5 (Yellow)', 'LED 7 (Yellow)', 'LED 9 (Orange)', 'LED 10 (Red - Trip)'];
            for (let i = 1; i <= 6; i++) {
                setVal(`load_kg_${i}`, `${tripLoads[i-1]}`);
                setVal(`input_mv_${i}`, `${((i - 1) * 4.0).toFixed(2)}`);
                setVal(`output_mv_${i}`, `${((i - 1) * 4.0).toFixed(2)}`);
                setVal(`led_level_${i}`, leds[i-1]);
            }
            setVal('trip_error', '0.02 %');
            setVal('pf_contact', 'Healthy Contact Verified OK');
            setVal('led_annunciation', 'Bar Graph Linear Response OK');
        } else if (reportType === 'signal_conditioner') {
            setVal('job_no', `JOB-SC-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('sc_model', 'SC-DIN-420');
            setVal('sc_serial', `SC-2026-${randJob}`);
            setVal('lc_model', 'LC-SHEAR-3000');
            setVal('lc_serial', `LC-88${randJob}`);
            setVal('cal_in_1', '0.0 mV'); setVal('cal_out_1', '4.00 mA');
            setVal('cal_in_2', '5.0 mV'); setVal('cal_out_2', '7.20 mA');
            setVal('cal_in_3', '10.0 mV'); setVal('cal_out_3', '10.40 mA');
            setVal('cal_in_4', '15.0 mV'); setVal('cal_out_4', '13.60 mA');
            setVal('cal_in_5', '20.0 mV'); setVal('cal_out_5', '16.80 mA');
            setVal('cal_in_6', '25.0 mV'); setVal('cal_out_6', '20.00 mA');
            setVal('annunciation', 'Checked & Found OK');
            setVal('pct_error', '0.02 %');
            setVal('multimeter_used', 'Fluke 87V (Sl No: FLK-9901)');
        } else if (reportType === 'vibration_switch') {
            setVal('job_no', `JOB-VS-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('model_no', 'VS-ELECTRONIC-100');
            setVal('serial_no', `SN-VS-88${randJob % 10}`);
            setVal('spec_1', '24V DC +/-10%'); setVal('act_1', '24.0V DC');
            setVal('spec_2', 'Velocity Trip: 4.5 mm/s'); setVal('act_2', '4.52 mm/s Passed');
            setVal('spec_3', 'Relay Delay: 3.0 sec'); setVal('act_3', '3.05 sec Passed');
            setVal('spec_4', '4-20mA Output Linear'); setVal('act_4', '4.0 - 20.0 mA Passed');
            setVal('cal_in_1', '0.0'); setVal('cal_out_1', '4.00 mA');
            setVal('cal_in_2', '1.5'); setVal('cal_out_2', '9.33 mA');
            setVal('cal_in_3', '3.0'); setVal('cal_out_3', '14.67 mA');
            setVal('cal_in_4', '4.5'); setVal('cal_out_4', '20.00 mA (Trip)');
        } else {
            // Belt Scale & Generic Defaults
            setVal('job_no', `JOB-2026-${randJob}`);
            setVal('customer', chosenClient);
            setVal('conveyor_no', `CV-${Math.floor(10 + Math.random() * 90)}B`);
            setVal('capacity', '1000');
            setVal('belt_speed', '1.75 m/s');
            setVal('tacho_belt_speed', '1.75');
            setVal('bs_model', `BS-MODEL-${randJob}`);
            setVal('bs_serial', `SN-BS-${7000 + randJob}`);
            setVal('remote_model', 'RM-UNIT-01');
            setVal('remote_serial', `SN-RM-${5000 + randJob}`);
            setVal('sensor_model', 'LC-250KG');
            setVal('s1', `LC1-00${randJob % 10}`);
            setVal('s2', `LC2-00${randJob % 10}`);
            setVal('s3', `LC3-00${randJob % 10}`);
            setVal('s4', `LC4-00${randJob % 10}`);
            setVal('s5', `LC5-00${randJob % 10}`);
            setVal('s6', `LC6-00${randJob % 10}`);
            setVal('s7', `LC7-00${randJob % 10}`);
            setVal('s8', `LC8-00${randJob % 10}`);
            setVal('tacho_model', 'TACHO-SENS-PRO');
            setVal('tacho_serial', `TS-${8000 + randJob}`);
            setVal('pulley_diameter', '0.16');
            setVal('tacho_wheel_dia', '0.318');
            setVal('tacho_rpm', '105.1 RPM');
            setVal('angle_sensor_model', 'AS-300-TILT');
            setVal('angle_sensor_serial', `SN-AS-${9000 + randJob}`);
            setVal('jbox_model1', 'JBOX-IP67'); setVal('jbox_serial1', `JB-SN-${100 + randJob}`);
            setVal('jbox_model2', 'JBOX-IP67'); setVal('jbox_serial2', `JB-SN-${200 + randJob}`);
            setVal('jbox_model3', 'JBOX-IP67'); setVal('jbox_serial3', `JB-SN-${300 + randJob}`);
            setVal('short_check', 'OK (Checked Line/Neutral, Neutral/Earth, Line/Earth)');
            setVal('res_l', '0.1'); setVal('res_s', '0.01'); setVal('res_r', '1'); setVal('res_t', '0.1');
            const demoFullLoad = (1000 * 1000.0) / (3600.0 * 1.75);
            for (let i = 1; i <= 9; i++) {
                setVal(`spec_${i}`, ['230V', '24V', 'OK', demoFullLoad.toFixed(1) + ' kg/m', '1.75 m/s', '4', 'RS-485', '4-20mA', 'OK'][i-1]);
                setVal(`act_${i}`, ['230V', '24V', 'OK', demoFullLoad.toFixed(1) + ' kg/m', '1.75 m/s', '4', 'RS-485', '4-20mA', 'OK'][i-1]);
            }
            for (let r = 1; r <= 6; r++) {
                const stepRatio = (r - 1) / 5.0;
                setVal(`g${r}_0`, (3.0 + stepRatio * 2.0).toFixed(2));
                setVal(`g${r}_1`, (stepRatio * demoFullLoad).toFixed(1));
                setVal(`g${r}_2`, '1.75');
                setVal(`g${r}_3`, (stepRatio * 1000).toFixed(0));
                setVal(`g${r}_4`, (stepRatio * 100).toFixed(1));
                setVal(`g${r}_5`, (4.0 + stepRatio * 16.0).toFixed(2));
                setVal(`g${r}_5_ao1`, (4.0 + stepRatio * 16.0).toFixed(2));
                setVal(`g${r}_5_ao2`, (4.0 + stepRatio * 16.0).toFixed(2));
                setVal(`g${r}_5_ao3`, (4.0 + stepRatio * 16.0).toFixed(2));
            }
        }

        // Common Sign-off & QA fields across all templates
        setVal('date', todayIso);
        setVal('tested_by', 'Field Inspection Engineer');
        setVal('approved_by', 'Quality Assurance HOD (PDN)');
        setVal('instrument_used', 'Fluke 87V Calibrated Multimeter & Class F Precision Weights');
        setVal('remarks', 'Routine testing and multi-point calibration completed successfully within specified engineering tolerances.');

        // 3. UNIVERSAL DEEP SWEEP: Fill EVERY remaining empty input, select, textarea, checkbox, radio in form or page
        const form = document.getElementById('recordForm') || document.querySelector('form');
        const inputs = (form || document).querySelectorAll('input, select, textarea');

        inputs.forEach(el => {
            if (!el || el.type === 'hidden' || el.type === 'submit' || el.type === 'button') return;

            // Checkbox: check all test/inspection pass checkboxes
            if (el.type === 'checkbox') {
                if (!el.checked) {
                    el.checked = true;
                    try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
                }
                return;
            }

            // Radio: check the first or the pass/ok option if none in group is checked
            if (el.type === 'radio') {
                const group = (form || document).querySelectorAll(`input[type="radio"][name="${el.name}"]`);
                const isAnyChecked = Array.from(group).some(r => r.checked);
                if (!isAnyChecked) {
                    const passRadio = Array.from(group).find(r => /pass|ok|yes|1|normal/i.test(r.value)) || group[0];
                    if (passRadio) {
                        passRadio.checked = true;
                        try { passRadio.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
                    }
                }
                return;
            }

            // Select: choose an active option if currently empty/first
            if (el.tagName === 'SELECT') {
                if (!el.value || el.selectedIndex === 0) {
                    let chosenIdx = 1;
                    for (let i = 0; i < el.options.length; i++) {
                        if (/pass|ok|normal|active|within|class\s*iii|230v|4-20ma/i.test(el.options[i].text || el.options[i].value)) {
                            chosenIdx = i;
                            break;
                        }
                    }
                    if (el.options.length > chosenIdx) {
                        el.selectedIndex = chosenIdx;
                        try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
                    }
                }
                return;
            }

            // Dates: set today's date if empty
            if (el.type === 'date') {
                if (!el.value) {
                    el.value = todayIso;
                    try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
                }
                return;
            }

            // Time: set current time if empty
            if (el.type === 'time') {
                if (!el.value) {
                    el.value = '10:00';
                    try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
                }
                return;
            }

            // Number: set sensible number if empty
            if (el.type === 'number') {
                if (!el.value) {
                    if (el.max && el.min) {
                        el.value = ((parseFloat(el.min) + parseFloat(el.max)) / 2).toFixed(el.step ? (el.step.includes('.') ? el.step.split('.')[1].length : 0) : 1);
                    } else if (el.step === '0.01' || el.step === '0.001') {
                        el.value = '10.00';
                    } else {
                        el.value = '100';
                    }
                    try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch(e){}
                    try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
                }
                return;
            }

            // Text & Textarea: if empty, fill realistic data matching field intent
            if (el.type === 'text' || el.tagName === 'TEXTAREA') {
                if (!el.value || el.value.trim() === '') {
                    const name = (el.name || el.id || '').toLowerCase();
                    const placeholder = (el.placeholder || '').toLowerCase();
                    let parentText = '';
                    try { parentText = (el.closest('.input-group, .form-group, td')?.innerText || '').toLowerCase(); } catch(e){}
                    const key = `${name} ${placeholder} ${parentText}`;

                    let fillVal = 'OK / Passed';
                    if (/ao\d*|output.*ma|current.*o/i.test(key) || /g\d+_5/i.test(name)) {
                        const mRow = name.match(/g(\d+)_/);
                        if (mRow) {
                            const rNum = parseInt(mRow[1]) || 1;
                            const stepRatio = (rNum - 1) / 5.0;
                            fillVal = (4.0 + stepRatio * 16.0).toFixed(2);
                        } else {
                            fillVal = '4.00';
                        }
                    }
                    else if (/spec_\d+|specified/i.test(name)) fillVal = '230V AC / Standard';
                    else if (/act_\d+|actual|measured/i.test(name)) fillVal = '230V AC / Passed';
                    else if (/job[_\s]*no|job[_\s]*num/i.test(key)) fillVal = `JOB-2026-${randJob}`;
                    else if (/cust/i.test(key)) fillVal = chosenClient;
                    else if (/model/i.test(key)) fillVal = 'MDL-PRO-2026';
                    else if (/serial|sr[_\s]*no/i.test(key)) fillVal = `SN-${randJob + 1000}`;
                    else if (/capacity|cap/i.test(key)) fillVal = '1000 TPH';
                    else if (/speed/i.test(key)) fillVal = '1.50 m/s';
                    else if (/res/i.test(key)) fillVal = '0.01';
                    else if (/volt/i.test(key)) fillVal = '230V AC';
                    else if (/curr|ma\b/i.test(key)) fillVal = '4.00 mA';
                    else if (/error|err/i.test(key)) fillVal = '0.02 %';
                    else if (/temp/i.test(key)) fillVal = '28 °C';
                    else if (/hum/i.test(key)) fillVal = '45 % RH';
                    else if (/press/i.test(key)) fillVal = '1.013 bar';
                    else if (/tested/i.test(key)) fillVal = 'Field Inspection Engineer';
                    else if (/appr/i.test(key)) fillVal = 'Quality Assurance HOD (PDN)';
                    else if (/remark|note|comment/i.test(key)) fillVal = 'All parameters verified and within specified engineering tolerances.';
                    else if (/inst|tool|equip/i.test(key)) fillVal = 'Fluke 87V Calibrated Multimeter & Precision Test Weights';

                    el.value = fillVal;
                    try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch(e){}
                    try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch(e){}
                }
            }
        });

        // 4. Trigger calculations across reactive formulas
        try { if (typeof calcRpm === 'function') calcRpm(); } catch(e){}
        try { if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix(); } catch(e){}
        try { if (typeof calculateMatrix === 'function') calculateMatrix(); } catch(e){}
        try { if (typeof recalculateAll === 'function') recalculateAll(); } catch(e){}

        // 5. Toast notification
        if (typeof showToast === 'function') {
            showToast('⚡ Auto-Fill Completed! Every field, table & checklist populated.', 'success');
        }
    }

    window.triggerUniversalAutoFill = triggerUniversalAutoFill;
    if (typeof window.fillDemoData !== 'function') {
        window.fillDemoData = triggerUniversalAutoFill;
    }

    // Ensure Auto-Fill button is present in top header across all report templates
    const headerActions = document.querySelector('.header-actions');
    if (headerActions && !document.getElementById('btnAutoFillHeader') && !window.location.pathname.includes('/login') && !window.location.pathname.includes('/portal')) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.id = 'btnAutoFillHeader';
        btn.className = 'header-nav-link btnFillDemoDataGlobal';
        btn.style.cssText = 'background: rgba(56, 189, 248, 0.16); border: 1px solid rgba(56, 189, 248, 0.4); color: #0284c7; cursor: pointer; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 8px;';
        btn.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles"></i> Auto-Fill Demo';
        headerActions.insertBefore(btn, headerActions.firstChild);
    }

    if (!window.__autoFillClickBound) {
        window.__autoFillClickBound = true;
        document.addEventListener('click', (e) => {
            const btn = e.target.closest('#btnAutoFillHeader, #btnFillSampleData, #btnTopAutoFill, .btnFillDemoDataGlobal, [data-action="autofill"]');
            if (btn) {
                e.preventDefault();
                triggerUniversalAutoFill();
            }
        });
    }

    // Tab Navigation Logic
    // ARCHIVE TOGGLE
    const btnToggleArchive = document.getElementById('btnToggleArchive');
    const btnCloseHistoryView = document.getElementById('btnCloseHistoryView');
    const generatorTab = document.getElementById('generatorTab');
    const historyTab = document.getElementById('historyTab');

    // INDUSTRIAL HUB NAVIGATION
    const hubView = document.getElementById('hubView');
    const btnLaunchBeltScale = document.getElementById('btnLaunchBeltScale');
    const cardModuleBeltScale = document.getElementById('cardModuleBeltScale');
    const btnBackToHub = document.getElementById('btnBackToHub');

    function openBeltScaleModule() {
        if (hubView) hubView.classList.remove('active');
        if (generatorTab) generatorTab.classList.add('active');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function backToHub() {
        if (generatorTab) generatorTab.classList.remove('active');
        if (hubView) hubView.classList.add('active');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    if (btnLaunchBeltScale) btnLaunchBeltScale.addEventListener('click', openBeltScaleModule);
    if (cardModuleBeltScale) cardModuleBeltScale.addEventListener('click', (e) => {
        if (e.target.id !== 'btnLaunchBeltScale') openBeltScaleModule();
    });
    if (btnBackToHub) btnBackToHub.addEventListener('click', backToHub);

    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const target = btn.getAttribute('data-tab');
            
            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));

            btn.classList.add('active');
            const targetElem = document.getElementById(target);
            if (targetElem) targetElem.classList.add('active');

            if (target === 'historyTab') {
                loadHistory();
            }
        });
    });

    // FORM SUBMISSION: ONLINE SERVER SUBMIT OR OFFLINE LOCALSTORAGE SAVE!
    const form = document.getElementById('recordForm');
    if (form) {
        function buildFormPayload() {
            const formData = new FormData(form);
            const payload = {};
            formData.forEach((value, key) => { payload[key] = value; });

            if (typeof window.extractRoutineTests === "function") {
                payload.routine_tests = window.extractRoutineTests();
            } else {
                payload.routine_tests = [];
                for (let i = 1; i <= 10; i++) {
                    const specElem = document.getElementById(`spec_${i}`) ||
                        document.querySelector(`[name="spec_${i}"]`) ||
                        document.getElementById(`t_spec_${i - 1}`);
                    const actElem = document.getElementById(`act_${i}`) ||
                        document.querySelector(`[name="act_${i}"]`) ||
                        document.getElementById(`t_act_${i - 1}`);
                    if (specElem || actElem || payload[`spec_${i}`] !== undefined || payload[`act_${i}`] !== undefined) {
                        const specVal = specElem ? specElem.value : (payload[`spec_${i}`] || '');
                        const actVal = actElem ? actElem.value : (payload[`act_${i}`] || '');
                        payload[`spec_${i}`] = specVal;
                        payload[`act_${i}`] = actVal;
                        payload.routine_tests.push({ spec: specVal, act: actVal });
                    }
                }
            }

            if (typeof window.extractGridData === "function") {
                payload.grid_data = window.extractGridData();
            } else {
                payload.grid_data = [];
                const hasDiCols = !!(document.getElementById('g1_6') || document.querySelector('[name="g1_6"]'))
                    && !(document.getElementById('g1_5_ao2') || document.querySelector('[name="g1_5_ao2"]'));
                const hasAoCols = !!(document.getElementById('g1_5_ao2') || document.querySelector('[name="g1_5_ao2"]'));
                const colCount = hasDiCols ? 10 : (hasAoCols ? 8 : 6);
                const maxRows = (document.getElementById('g12_0') || document.querySelector('[name="g12_0"]')) ? 12 : 11;

                for (let r = 1; r <= maxRows; r++) {
                    const rowVals = [];
                    for (let c = 0; c < colCount; c++) {
                        let elem = null;
                        if (hasAoCols && c === 5) {
                            elem = document.getElementById(`g${r}_5_ao1`) || document.querySelector(`[name="g${r}_5_ao1"]`) || document.getElementById(`g${r}_5`);
                        } else if (hasAoCols && c === 6) {
                            elem = document.getElementById(`g${r}_5_ao2`) || document.querySelector(`[name="g${r}_5_ao2"]`);
                        } else if (hasAoCols && c === 7) {
                            elem = document.getElementById(`g${r}_5_ao3`) || document.querySelector(`[name="g${r}_5_ao3"]`);
                        } else {
                            elem = document.getElementById(`g${r}_${c}`) || document.getElementById(`g_${r - 1}_${c}`) || document.querySelector(`[name="g${r}_${c}"]`);
                        }
                        const cellVal = elem ? elem.value : (payload[`g${r}_${c}`] || '');
                        rowVals.push(cellVal);
                    }
                    if (rowVals.length > 0) {
                        payload.grid_data.push(rowVals);
                    }
                }
            }
            payload.timestamp = new Date().toISOString();
            return payload;
        }

        function openPdfPreviewModal(pdfUrl, pdfFilename, pdfGenerated, docxUrl, docxFilename) {
            if (typeof window.showPdfPreview === "function" && pdfGenerated !== false && pdfUrl) {
                window.showPdfPreview(pdfUrl, pdfFilename);
                return;
            }
            const modal = document.getElementById("pdfPreviewModal");
            if (!modal) return;

            const titleEl = document.getElementById("modalPreviewTitle");
            const iframe = document.getElementById("pdfPreviewIframe");
            const dlBtn = document.getElementById("btnModalDownloadPDF");
            let msg = document.getElementById('androidFixMsg');

            const isDocxOnly = (pdfGenerated === false) || (!pdfUrl && !!docxUrl);
            const activeUrl = isDocxOnly ? docxUrl : pdfUrl;
            const activeName = isDocxOnly ? (docxFilename || 'Report.docx') : (pdfFilename || 'Report.pdf');
            if (titleEl) titleEl.textContent = activeName;
            if (dlBtn && activeUrl) {
                dlBtn.href = `${activeUrl}${activeUrl.includes('?') ? '&' : '?'}download=1`;
                dlBtn.setAttribute('download', activeName);
            }

            if (iframe) {
                if (/Android/i.test(navigator.userAgent) || isDocxOnly || activeName.toLowerCase().endsWith('.docx')) {
                    iframe.style.display = 'none';
                    iframe.removeAttribute('src');
                    if (!msg) {
                        msg = document.createElement('div');
                        msg.id = 'androidFixMsg';
                        msg.style.padding = '40px';
                        msg.style.textAlign = 'center';
                        msg.style.color = '#fff';
                        iframe.parentNode.appendChild(msg);
                    }
                    msg.style.display = 'block';
                    msg.innerHTML = `<i class="fa-solid fa-file-word" style="font-size:48px; color:#38bdf8; margin-bottom:16px;"></i><br><h3>Document Generated Successfully!</h3><p>PDF inline preview is unavailable here. Use <b>Download</b> above to save <code>${escapeHtml(activeName)}</code>.</p>`;
                } else {
                    if (msg) msg.style.display = 'none';
                    iframe.style.display = 'block';
                    iframe.src = `${activeUrl}${activeUrl.includes('?') ? '&' : '?'}inline=1`;
                }
            }
            modal.style.display = "flex";
        }

        function openDryPreviewModal(payload) {
            let existing = document.getElementById('dryFormPreviewModal');
            if (existing) existing.remove();

            const modal = document.createElement('div');
            modal.id = 'dryFormPreviewModal';
            modal.className = 'modal-overlay';
            modal.style.cssText = 'display:flex; position:fixed; inset:0; background:rgba(15,23,42,0.85); backdrop-filter:blur(6px); z-index:9999; justify-content:center; align-items:center; padding:16px;';

            const entries = Object.entries(payload).filter(([k, v]) =>
                !['grid_data', 'routine_tests', 'timestamp', 'report_type'].includes(k) &&
                v !== undefined && v !== null && String(v).trim() !== ''
            );
            const rowsHtml = entries.map(([k, v]) =>
                `<tr><td style="font-weight:700; color:#334155; padding:8px 12px; border-bottom:1px solid #e2e8f0; width:38%;">${escapeHtml(k)}</td><td style="padding:8px 12px; border-bottom:1px solid #e2e8f0; color:#0f172a;">${escapeHtml(String(v))}</td></tr>`
            ).join('');

            modal.innerHTML = `
                <div style="background:#fff; border-radius:16px; width:min(760px, 95vw); max-height:85vh; display:flex; flex-direction:column; overflow:hidden; box-shadow:0 24px 60px rgba(0,0,0,0.4);">
                    <div style="padding:16px 20px; background:#0f172a; color:#fff; display:flex; align-items:center; justify-content:space-between;">
                        <h3 style="margin:0; font-size:16px;"><i class="fa-solid fa-eye"></i> Draft Data Preview (Not Yet Saved)</h3>
                        <button type="button" id="btnCloseDryPreview" style="background:transparent; border:none; color:#94a3b8; font-size:20px; cursor:pointer;">&times;</button>
                    </div>
                    <div style="overflow:auto; padding:16px 20px; flex:1;">
                        <table style="width:100%; border-collapse:collapse; font-size:13.5px;">${rowsHtml || '<tr><td style="padding:16px; color:#64748b;">No fields entered yet.</td></tr>'}</table>
                    </div>
                    <div style="padding:14px 20px; background:#f8fafc; border-top:1px solid #e2e8f0; display:flex; justify-content:flex-end; gap:10px;">
                        <button type="button" id="btnCancelDryPreview" class="btn-secondary">Back to Edit</button>
                        <button type="button" id="btnConfirmDrySubmit" class="btn-primary"><i class="fa-solid fa-file-export"></i> Generate & Save Official Report</button>
                    </div>
                </div>
            `;
            document.body.appendChild(modal);

            const close = () => modal.remove();
            modal.querySelector('#btnCloseDryPreview').onclick = close;
            modal.querySelector('#btnCancelDryPreview').onclick = close;
            modal.addEventListener('click', (e) => { if (e.target === modal) close(); });
            modal.querySelector('#btnConfirmDrySubmit').onclick = () => {
                close();
                submitFormHandler();
            };
        }

        const btnPreview = document.getElementById('btnPreviewForm');
        if (btnPreview) {
            btnPreview.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopImmediatePropagation();
                openDryPreviewModal(buildFormPayload());
            });
        }

        const origSubmitBtnHtml = btnSubmitPDF ? btnSubmitPDF.innerHTML : '<i class="fa-solid fa-file-export"></i> Submit Report';

        const submitFormHandler = async (e) => {
            if (e) e.preventDefault();
            if (form && form.dataset.submitting === "true") return;
            if (form) form.dataset.submitting = "true";

            if (btnSubmitPDF) {
                btnSubmitPDF.disabled = true;
                btnSubmitPDF.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processing Document...';
            }

            let payload = null;
            try {
                payload = buildFormPayload();
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 30000);

                const response = await fetch(`${getServerBaseUrl()}/api/submit`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload),
                    signal: controller.signal
                });
                clearTimeout(timeoutId);

                if (response.status === 401 || (response.redirected && response.url.includes('/login'))) {
                    showToast('Session expired. Please sign in again before submitting.', 'warning');
                    return;
                }

                const contentType = response.headers.get('content-type') || '';
                if (!contentType.includes('application/json')) {
                    showToast(`Server returned unexpected status ${response.status}.`, 'error');
                    return;
                }
                const result = await response.json();

                if (response.ok && result.status === 'success') {
                    isServerOnline = true;
                    updateNetworkStatusIndicator();
                    showToast('Report Generated Successfully!', 'success');

                    const docxUrl = result.docx_download_url || (result.docx_filename ? `${getServerBaseUrl()}/download/${encodeURIComponent(result.docx_filename)}` : null);
                    const docxFilename = result.docx_filename || "Report.docx";

                    const pdfUrl = (result.pdf_generated !== false && (result.pdf_download_url || result.pdf_filename))
                        ? (result.pdf_download_url || `${getServerBaseUrl()}/download/${encodeURIComponent(result.pdf_filename)}`)
                        : null;
                    const pdfFilename = result.pdf_filename || (docxFilename ? docxFilename.replace(/\.docx$/i, '.pdf') : "Report.pdf");

                    const mainDlUrl = pdfUrl || docxUrl;
                    const mainFilename = pdfUrl ? pdfFilename : docxFilename;

                    if (mainDlUrl) {
                        const banner = document.getElementById("downloadBanner");
                        const bannerFile = document.getElementById("downloadFilename");
                        const btnDlPDF = document.getElementById("btnDownloadPDF");
                        const btnDlDOCX = document.getElementById("btnDownloadDOCX");
                        const btnBannerPreview = document.getElementById("btnBannerPreviewPDF");

                        if (banner) {
                            banner.classList.add("show-banner");
                            banner.style.setProperty("display", "flex", "important");
                            if (bannerFile) bannerFile.textContent = mainFilename;

                            let closeBtn = banner.querySelector(".btn-banner-close");
                            if (!closeBtn) {
                                closeBtn = document.createElement("button");
                                closeBtn.type = "button";
                                closeBtn.className = "btn-banner-close";
                                closeBtn.innerHTML = '<i class="fa-solid fa-xmark"></i>';
                                closeBtn.title = "Dismiss";
                                banner.appendChild(closeBtn);
                            }
                            closeBtn.onclick = () => window.closeDownloadBanner();

                            if (btnDlPDF) {
                                if (pdfUrl) {
                                    btnDlPDF.style.display = '';
                                    btnDlPDF.href = `${pdfUrl}${pdfUrl.includes('?') ? '&' : '?'}download=1`;
                                    btnDlPDF.setAttribute("download", pdfFilename);
                                } else {
                                    btnDlPDF.style.display = 'none';
                                }
                            }
                            if (btnDlDOCX && docxUrl) {
                                btnDlDOCX.href = `${docxUrl}${docxUrl.includes('?') ? '&' : '?'}download=1`;
                                btnDlDOCX.setAttribute("download", docxFilename);
                            }
                            if (btnBannerPreview) {
                                btnBannerPreview.onclick = (ev) => {
                                    ev.preventDefault();
                                    openPdfPreviewModal(pdfUrl, pdfFilename, result.pdf_generated, docxUrl, docxFilename);
                                };
                            }
                            try {
                                banner.scrollIntoView({ behavior: "smooth", block: "center" });
                            } catch (e) {}
                        }

                        try {
                            if (docxUrl) {
                                downloadWithSaveAsPrompt(`${docxUrl}${docxUrl.includes('?') ? '&' : '?'}download=1`, docxFilename);
                            }
                        } catch (e) { console.error("DOCX Auto-download error:", e); }

                        try {
                            if (pdfUrl) {
                                setTimeout(() => {
                                    try { downloadWithSaveAsPrompt(`${pdfUrl}${pdfUrl.includes('?') ? '&' : '?'}download=1`, pdfFilename); } catch (e) {}
                                }, 600);
                            }
                        } catch (e) { console.error("PDF Auto-download error:", e); }

                        showToast('Auto-Downloading generated report files...', 'success');
                    }

                    await loadHistory();
                    return;
                } else {
                    showToast('Submit Error: ' + (result.message || 'Unknown server error'), 'error');
                }
            } catch (err) {
                console.error('Submit error:', err);
                const isNetworkErr = !navigator.onLine || err.name === 'TypeError' || err.name === 'AbortError';
                if (isNetworkErr && payload) {
                    isServerOnline = false;
                    updateNetworkStatusIndicator();
                    saveOfflineRecord(payload);
                } else {
                    showToast('Submit Error: ' + (err.message || err), 'error');
                }
            } finally {
                if (form) form.dataset.submitting = "false";
                if (btnSubmitPDF) {
                    btnSubmitPDF.disabled = false;
                    btnSubmitPDF.innerHTML = origSubmitBtnHtml;
                }
            }
        };

        if (form) {
            form.addEventListener('submit', submitFormHandler);
        }
        document.addEventListener('click', (e) => {
            const btn = e.target.closest('#btnSubmitPDF, .btn-submit-report');
            if (btn && !btn.closest('#addUserForm') && !btn.closest('#loginForm')) {
                e.preventDefault();
                submitFormHandler(e);
            }
        });
        window.executeBeltScaleSubmit = submitFormHandler;
        window.submitBeltScaleForm = submitFormHandler;
        window.submitReportForm = submitFormHandler;
    }

    // Load History Function (Server Records + Local Offline Queue)
    async function loadHistory() {
        const historyBody = document.getElementById('historyTableBody');
        if (!historyBody) return;

        let serverRecords = [];
        try {
            const res = await fetch(`${getServerBaseUrl()}/api/records`);
            const data = await res.json();
            if (data.status === 'success' && data.records) {
                serverRecords = data.records;
            }
        } catch (e) {}

        const offlineRecords = getOfflineRecords();
        updateOfflineBadge();

        if (serverRecords.length === 0 && offlineRecords.length === 0) {
            historyBody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding: 24px; color:var(--text-muted);">No PDF documents created yet. Fill data and click Save & Generate PDF!</td></tr>';
            return;
        }

        let html = '';

        // Render Pending Offline Records first
        if (offlineRecords.length > 0) {
            html += offlineRecords.map(r => `
                <tr style="background: rgba(245, 158, 11, 0.12); border-left: 4px solid #f59e0b;">
                    <td><code style="background:rgba(245, 158, 11, 0.25); color:#fbbf24; padding:4px 8px; border-radius:6px; font-weight:800;">${escapeHtml(r.job_no || '-')}</code></td>
                    <td><strong>${escapeHtml(r.customer || '-')}</strong></td>
                    <td>${escapeHtml(r.conveyor_no || '-')}</td>
                    <td><i class="fa-solid fa-user-circle" style="color:#fbbf24;"></i> ${escapeHtml(r.user_name || '-')}</td>
                    <td><span style="color:#fbbf24; font-weight:700;">⚡ Offline (${escapeHtml(r._created_at || '')})</span></td>
                    <td style="text-align: center;">
                        <span style="background:rgba(245, 158, 11, 0.2); color:#fbbf24; border:1px solid rgba(245, 158, 11, 0.5); padding:4px 12px; border-radius:14px; font-size:12px; font-weight:700;">
                            Pending Server Sync
                        </span>
                    </td>
                </tr>
            `).join('');
        }

        // Render Server Records
        if (serverRecords.length > 0) {
            const baseUrl = getServerBaseUrl();
            html += serverRecords.map(r => {
                const pdfName = r.pdf_filename || (r.docx_filename ? r.docx_filename.replace(/\.docx$/i, '.pdf') : '');
                const docName = r.docx_filename || '';
                const viewUrl = pdfName ? `${baseUrl}/download/${encodeURIComponent(pdfName)}?inline=1` : '';
                const dlName = pdfName || docName;
                const dlUrl = dlName ? `${baseUrl}/download/${encodeURIComponent(dlName)}?download=1` : '';
                const actionsHtml = dlName
                    ? `${viewUrl ? `<a href="${escapeHtml(viewUrl)}" target="_blank" class="btn-secondary" style="font-size:12px; padding:6px 12px; border-color:rgba(56, 189, 248, 0.4); color:#38bdf8;"><i class="fa-solid fa-eye"></i> View PDF</a>` : ''}
                       <a href="${escapeHtml(dlUrl)}" class="btn-download" download><i class="fa-solid fa-download"></i> Download</a>`
                    : `<span style="color:var(--text-muted); font-size:12px;">No file</span>`;
                return `
                <tr>
                    <td><code style="background:rgba(56, 189, 248, 0.15); color:#38bdf8; padding:4px 8px; border-radius:6px; font-weight:800;">${escapeHtml(r.job_no || '-')}</code></td>
                    <td><strong>${escapeHtml(r.customer || '-')}</strong></td>
                    <td>${escapeHtml(r.conveyor_no || '-')}</td>
                    <td><i class="fa-solid fa-user-circle" style="color:#38bdf8;"></i> ${escapeHtml(r.user_name || '-')}</td>
                    <td>${escapeHtml(r.created_at || '-')}</td>
                    <td style="text-align: center; display: flex; gap: 8px; justify-content: center;">
                        ${actionsHtml}
                    </td>
                </tr>`;
            }).join('');
        }

        historyBody.innerHTML = html;
    }

    const btnRefreshHistoryList = document.getElementById('btnRefreshHistoryList');
    if (btnRefreshHistoryList) btnRefreshHistoryList.addEventListener('click', loadHistory);

    loadHistory();

    // Toast Function
    function showToast(msg, type = 'success') {
        const toast = document.getElementById('toast');
        const toastMsg = document.getElementById('toastMsg');
        const toastIcon = document.getElementById('toastIcon');

        if (!toast || !toastMsg) return;

        toastMsg.innerText = msg;
        if (type === 'error') {
            toastIcon.className = 'fa-solid fa-circle-exclamation';
            toast.style.borderColor = '#f43f5e';
        } else if (type === 'warning') {
            toastIcon.className = 'fa-solid fa-triangle-exclamation';
            toast.style.borderColor = '#f59e0b';
        } else {
            toastIcon.className = 'fa-solid fa-circle-check';
            toast.style.borderColor = '#10b981';
        }

        toast.classList.remove('hidden');
        setTimeout(() => {
            toast.classList.add('hidden');
        }, 4000);
    }

    // INTERACTIVE REGION SELECTION SNIPPING TOOL & RULES LOGGING ENGINE
    function initScreenshotCapture() {
        if (!window.html2canvas) {
            const script = document.createElement('script');
            script.src = '/static/js/html2canvas.min.js';
            document.head.appendChild(script);
        }

        const floatBtn = document.createElement('button');
        floatBtn.id = 'btnFloatingScreenshot';
        floatBtn.title = 'Select Region Screenshot & Log Rule (Ctrl+Shift+S or Ctrl+Shift+I)';
        floatBtn.innerHTML = '<i class="fa-solid fa-crop-simple"></i>';
        floatBtn.style.cssText = `
            position: fixed;
            bottom: 80px;
            right: 24px;
            width: 52px;
            height: 52px;
            border-radius: 50%;
            background: linear-gradient(135deg, #0284c7, #0ea5e9);
            color: #ffffff;
            border: 2px solid rgba(255, 255, 255, 0.4);
            font-size: 20px;
            cursor: pointer;
            box-shadow: 0 8px 24px rgba(2, 132, 199, 0.4);
            z-index: 9999;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
        `;
        floatBtn.onmouseover = () => floatBtn.style.transform = 'scale(1.1)';
        floatBtn.onmouseout = () => floatBtn.style.transform = 'scale(1.0)';
        floatBtn.onclick = startRegionSnipping;
        document.body.appendChild(floatBtn);

        document.addEventListener('keydown', (e) => {
            if ((e.ctrlKey && e.shiftKey && (e.key === 'S' || e.key === 's')) || 
                (e.ctrlKey && e.shiftKey && (e.key === 'I' || e.key === 'i'))) {
                e.preventDefault();
                startRegionSnipping();
            }
        });

        async function captureFullScreen() {
            try {
                showToast('Capturing screen...', 'warning');
                floatBtn.style.display = 'none';
                const canvas = await html2canvas(document.body, {
                    useCORS: true,
                    logging: false,
                    scale: window.devicePixelRatio || 1
                });
                floatBtn.style.display = 'flex';
                const imageData = canvas.toDataURL('image/png');
                showScreenshotModal(imageData);
            } catch(err) {
                floatBtn.style.display = 'flex';
                showToast('Failed to capture screen: ' + err.message, 'error');
            }
        }

        function startRegionSnipping() {
            if (!window.html2canvas) {
                showToast('Loading snipping engine...', 'warning');
                setTimeout(startRegionSnipping, 800);
                return;
            }

            let existingOverlay = document.getElementById('snippingOverlay');
            if (existingOverlay) existingOverlay.remove();

            const overlay = document.createElement('div');
            overlay.id = 'snippingOverlay';
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; right: 0; bottom: 0;
                width: 100vw; height: 100vh; background: rgba(15, 23, 42, 0.35);
                cursor: crosshair; z-index: 100000; user-select: none;
                touch-action: none;
            `;

            const banner = document.createElement('div');
            banner.style.cssText = `
                position: absolute; top: 20px; left: 50%; transform: translateX(-50%);
                background: #1e293b; color: #38bdf8; border: 1.5px solid rgba(56, 189, 248, 0.5);
                padding: 12px 28px; border-radius: 30px; font-family: 'Outfit', sans-serif;
                font-size: 14px; font-weight: 700; box-shadow: 0 8px 24px rgba(0,0,0,0.6);
                pointer-events: none; display: flex; align-items: center; gap: 10px;
                max-width: 90vw; text-align: center;
            `;
            banner.innerHTML = '<i class="fa-solid fa-crop-simple"></i> Drag region or tap screen to capture';
            overlay.appendChild(banner);

            const selectionBox = document.createElement('div');
            selectionBox.style.cssText = `
                position: absolute; border: 2px dashed #0ea5e9; background: rgba(14, 165, 233, 0.2);
                box-shadow: 0 0 0 9999px rgba(15, 23, 42, 0.45); display: none; pointer-events: none;
            `;
            overlay.appendChild(selectionBox);

            let startX = 0, startY = 0, isDragging = false;

            const handleStart = (clientX, clientY) => {
                startX = clientX;
                startY = clientY;
                isDragging = true;
                selectionBox.style.left = `${startX}px`;
                selectionBox.style.top = `${startY}px`;
                selectionBox.style.width = '0px';
                selectionBox.style.height = '0px';
                selectionBox.style.display = 'block';
            };

            const handleMove = (clientX, clientY) => {
                if (!isDragging) return;
                const rectLeft = Math.min(startX, clientX);
                const rectTop = Math.min(startY, clientY);
                const rectWidth = Math.abs(clientX - startX);
                const rectHeight = Math.abs(clientY - startY);

                selectionBox.style.left = `${rectLeft}px`;
                selectionBox.style.top = `${rectTop}px`;
                selectionBox.style.width = `${rectWidth}px`;
                selectionBox.style.height = `${rectHeight}px`;
            };

            overlay.onmousedown = (e) => handleStart(e.clientX, e.clientY);
            overlay.onmousemove = (e) => handleMove(e.clientX, e.clientY);
            
            overlay.ontouchstart = (e) => {
                if (e.touches && e.touches.length > 0) {
                    handleStart(e.touches[0].clientX, e.touches[0].clientY);
                }
            };
            overlay.ontouchmove = (e) => {
                if (e.touches && e.touches.length > 0) {
                    handleMove(e.touches[0].clientX, e.touches[0].clientY);
                }
            };

            const cancelSnipping = (e) => {
                if (e.key === 'Escape') {
                    overlay.remove();
                    document.removeEventListener('keydown', cancelSnipping);
                }
            };
            document.addEventListener('keydown', cancelSnipping);

            overlay.onmouseup = (e) => endDrag(e.clientX, e.clientY);
            overlay.ontouchend = (e) => {
                if (e.changedTouches && e.changedTouches.length > 0) {
                    endDrag(e.changedTouches[0].clientX, e.changedTouches[0].clientY);
                } else {
                    endDrag(startX, startY);
                }
            };

            const endDrag = async (clientX, clientY) => {
                if (!isDragging) return;
                isDragging = false;
                document.removeEventListener('keydown', cancelSnipping);

                const currentX = clientX;
                const currentY = clientY;
                const rectLeft = Math.min(startX, currentX);
                const rectTop = Math.min(startY, currentY);
                const rectWidth = Math.abs(currentX - startX);
                const rectHeight = Math.abs(currentY - startY);

                overlay.remove();

                if (rectWidth < 15 || rectHeight < 15) {
                    captureFullScreen();
                    return;
                }

                floatBtn.style.display = 'none';

                try {
                    showToast('Capturing selected region...', 'warning');
                    
                    const fullCanvas = await html2canvas(document.body, {
                        useCORS: true,
                        logging: false,
                        scale: window.devicePixelRatio || 1
                    });
                    
                    floatBtn.style.display = 'flex';

                    const scale = window.devicePixelRatio || 1;
                    const cropCanvas = document.createElement('canvas');
                    cropCanvas.width = rectWidth * scale;
                    cropCanvas.height = rectHeight * scale;
                    const ctx = cropCanvas.getContext('2d');

                    const scrollX = window.scrollX || window.pageXOffset;
                    const scrollY = window.scrollY || window.pageYOffset;

                    ctx.drawImage(
                        fullCanvas,
                        (rectLeft + scrollX) * scale,
                        (rectTop + scrollY) * scale,
                        rectWidth * scale,
                        rectHeight * scale,
                        0,
                        0,
                        rectWidth * scale,
                        rectHeight * scale
                    );

                    const croppedImageData = cropCanvas.toDataURL('image/png');
                    showScreenshotModal(croppedImageData);

                } catch(err) {
                    floatBtn.style.display = 'flex';
                    showToast('Failed to capture region: ' + err.message, 'error');
                }
            };

            document.body.appendChild(overlay);
        }

        function showScreenshotModal(imageData) {
            let modal = document.getElementById('screenshotModal');
            if (modal) modal.remove();

            modal = document.createElement('div');
            modal.id = 'screenshotModal';
            modal.style.cssText = `
                position: fixed; top: 0; left: 0; right: 0; bottom: 0;
                background: rgba(15, 23, 42, 0.85); backdrop-filter: blur(8px);
                z-index: 100000; display: flex; align-items: center; justify-content: center; padding: 20px;
            `;

            modal.innerHTML = `
                <div style="background: #1e293b; border: 1px solid rgba(56, 189, 248, 0.3); border-radius: 20px; max-width: 580px; width: 100%; padding: 28px; box-shadow: 0 20px 40px rgba(0,0,0,0.5); color: #ffffff;">
                    <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;">
                        <h3 style="font-family: 'Outfit', sans-serif; font-size: 20px; font-weight: 700; color: #38bdf8; display: flex; align-items: center; gap: 10px; margin: 0;">
                            <i class="fa-solid fa-crop-simple"></i> Selected Region Screenshot & Rule Note
                        </h3>
                        <button id="closeScreenshotModal" style="background: none; border: none; color: #94a3b8; font-size: 20px; cursor: pointer;">&times;</button>
                    </div>

                    <div style="margin-bottom: 20px; text-align: center; background: #0f172a; padding: 12px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.1);">
                        <img src="${imageData}" style="max-width: 100%; max-height: 240px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.3);" />
                    </div>

                    <div style="margin-bottom: 20px;">
                        <label style="display: block; font-size: 13px; font-weight: 700; margin-bottom: 8px; color: #cbd5e1;">Rule Note / Feedback Directive:</label>
                        <textarea id="screenshotNoteInput" placeholder="Enter rule note or UI feedback directive for this region..." style="width: 100%; height: 80px; background: #0f172a; border: 1px solid rgba(255,255,255,0.15); border-radius: 10px; padding: 12px; color: #ffffff; font-family: inherit; font-size: 14px; resize: vertical; box-sizing: border-box;"></textarea>
                    </div>

                    <div style="display: flex; gap: 12px; justify-content: flex-end;">
                        <button id="cancelScreenshotModal" style="background: #334155; color: #ffffff; border: none; padding: 10px 20px; border-radius: 10px; font-weight: 700; cursor: pointer;">Cancel</button>
                        <button id="saveScreenshotModal" style="background: linear-gradient(135deg, #10b981, #059669); color: #ffffff; border: none; padding: 10px 22px; border-radius: 10px; font-weight: 700; cursor: pointer; box-shadow: 0 4px 14px rgba(16,185,129,0.3);">
                            <i class="fa-solid fa-cloud-arrow-up"></i> Save & Log Rule
                        </button>
                    </div>
                </div>
            `;

            document.body.appendChild(modal);

            document.getElementById('closeScreenshotModal').onclick = () => modal.remove();
            document.getElementById('cancelScreenshotModal').onclick = () => modal.remove();

            document.getElementById('saveScreenshotModal').onclick = async () => {
                const note = document.getElementById('screenshotNoteInput').value;
                const saveBtn = document.getElementById('saveScreenshotModal');
                saveBtn.disabled = true;
                saveBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Saving...';

                try {
                    const res = await fetch(`${getServerBaseUrl()}/api/capture_screenshot`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            image_data: imageData,
                            note: note,
                            page_url: window.location.pathname
                        })
                    });
                    const data = await res.json();
                    if (data.status === 'success') {
                        modal.remove();
                        showToast('Region saved to LIVE TESTING/ & SCREENSHOTS/ and logged to RULES.md!', 'success');
                    } else {
                        showToast(data.message || 'Failed to save region screenshot', 'error');
                        saveBtn.disabled = false;
                        saveBtn.innerHTML = '<i class="fa-solid fa-cloud-arrow-up"></i> Save & Log Rule';
                    }
                } catch(e) {
                    showToast('Error saving screenshot to server', 'error');
                    saveBtn.disabled = false;
                    saveBtn.innerHTML = '<i class="fa-solid fa-cloud-arrow-up"></i> Save & Log Rule';
                }
            };
        }
    }

    initScreenshotCapture();
});

// ====== CUSTOM BUSINESS RULES ENGINE ======
document.addEventListener('DOMContentLoaded', () => {
    const path = window.location.pathname.toLowerCase();
    const normPath = path.replace(/-/g, '_');

    const num = (val) => {
        const cleaned = String(val === undefined || val === null ? '' : val).replace(/,/g, '');
        const m = cleaned.match(/-?\d+(?:\.\d+)?/);
        return m ? parseFloat(m[0]) : 0;
    };

    // 1. Weigh Feeder / Screwfeeder / Loss In Weigh Feeder
    if (normPath.includes('weigh_feeder') || normPath.includes('screwfeeder')) {
        const applyWeighFeederRules = () => {
            const speedEl = document.getElementById('belt_speed') || document.querySelector('[name="belt_speed"]');
            const rateEl = document.getElementById('capacity') || document.querySelector('[name="capacity"]');
            
            let s = 1.0;
            if (speedEl && speedEl.value) {
                let speedVal = String(speedEl.value).toLowerCase();
                s = num(speedVal);
                if (speedVal.includes('min')) s = s / 60.0;
                
                let sDec = s >= 0.1 ? 3 : 4;
                speedEl.value = s.toFixed(sDec) + " m/s";
                
                const resSpeedEl = document.getElementById('res_speed') || document.querySelector('[name="res_speed"]');
                if (resSpeedEl) resSpeedEl.value = (1 / Math.pow(10, sDec)).toFixed(sDec) + " m/s";
            }
            
            if (rateEl && rateEl.value) {
                let r = num(rateEl.value);
                let rDec = 2;
                if (r >= 100 && r <= 999) rDec = 1;
                if (r > 999) rDec = 0;
                rateEl.value = r.toFixed(rDec) + (String(rateEl.value).toLowerCase().includes('tph') ? ' TPH' : '');
                
                const resRateEl = document.getElementById('res_rate') || document.querySelector('[name="res_rate"]');
                if (resRateEl) resRateEl.value = (1 / Math.pow(10, rDec)).toFixed(rDec) + " tph";
                
                // belt load
                if (s > 0) {
                    let bl = r / (3.6 * s);
                    let blDec = 3;
                    if (bl >= 10) blDec = 2;
                    
                    const resLoadEl = document.getElementById('res_load') || document.querySelector('[name="res_load"]');
                    if (resLoadEl) resLoadEl.value = (1 / Math.pow(10, blDec)).toFixed(blDec) + " kg/m";
                    
                    const beltLoadEl = document.getElementById('belt_load') || document.querySelector('[name="belt_load"]');
                    if (beltLoadEl) beltLoadEl.value = bl.toFixed(blDec) + " kg/m";
                }
                
                // totaliser
                const resTotEl = document.getElementById('res_tot') || document.querySelector('[name="res_tot"]') || document.querySelector('[name="totaliser"]');
                if (resTotEl) resTotEl.value = (r / 10).toFixed(rDec);
            }
        };

        // Removed unwanted periodic setInterval. Rules trigger only on user input/change:
        document.body.addEventListener('change', (e) => {
            if (e.target.name === 'belt_speed' || e.target.name === 'capacity') {
                applyWeighFeederRules();
            }
        });
    }

    // 2. Belt Scale
    if (normPath.includes('belt_scale')) {
        const applyBeltScaleRules = () => {
            const speedEl = document.getElementById('belt_speed') || document.querySelector('[name="belt_speed"]');
            const rateEl = document.getElementById('capacity') || document.querySelector('[name="capacity"]');
            if (!speedEl || !rateEl || !speedEl.value || !rateEl.value) return;
            
            let speedNum = num(speedEl.value);
            let rateNum = num(rateEl.value);
            if (speedNum <= 0 || rateNum <= 0) return;

            const calcResRule = (v) => {
                const num = Math.abs(parseFloat(v) || 0);
                if (num <= 99) return { str: '0.01', dec: 2 };
                if (num <= 999) return { str: '0.1', dec: 1 };
                return { str: '1', dec: 0 };
            };

            let fullLoad = (rateNum * 1000.0) / (3600.0 * speedNum);
            let resL = calcResRule(fullLoad);
            let resR = calcResRule(rateNum);
            let resS = { str: '0.01', dec: 2 };
            let resT = { str: '0.1', dec: 1 };

            const resLEl = document.getElementById('res_l') || document.querySelector('[name="res_l"]');
            if (resLEl && !resLEl.value) resLEl.value = resL.str;
            const resREl = document.getElementById('res_r') || document.querySelector('[name="res_r"]');
            if (resREl && !resREl.value) resREl.value = resR.str;
            const resSEl = document.getElementById('res_s') || document.querySelector('[name="res_s"]');
            if (resSEl && !resSEl.value) resSEl.value = resS.str;
            const resTEl = document.getElementById('res_t') || document.querySelector('[name="res_t"]');
            if (resTEl && !resTEl.value) resTEl.value = resT.str;

            const beltLoadEl = document.getElementById('belt_load') || document.querySelector('[name="belt_load"]');
            if (beltLoadEl && !beltLoadEl.value) beltLoadEl.value = fullLoad.toFixed(resL.dec) + " kg/m";
        };

        // Removed periodic setInterval. Trigger only on user input/change:
        document.body.addEventListener('change', (e) => {
            if (e.target.name === 'belt_speed' || e.target.name === 'capacity') {
                applyBeltScaleRules();
            }
        });
    }

    // 3. CWS (Crane Scale)
    if (normPath.includes('crane_scale') || normPath.includes('cws')) {
        const calcMvSteps = () => {
            const finalMvEl = document.getElementById('final_mv_output') || document.querySelector('[name="final_mv_output"]');
            const lcEl = document.getElementById('capacity') || document.querySelector('[name="capacity"]');
            if (!finalMvEl || !finalMvEl.value) return;
            
            let finalMv = num(finalMvEl.value);
            let lc = num(lcEl ? lcEl.value : "100");
            
            let steps = [0.0, 0.25, 0.50, 0.75, 1.0];
            steps.forEach((pct, idx) => {
                let stepMv = (finalMv * pct).toFixed(3);
                let stepLoad = (lc * pct).toFixed(1);
                
                let mvField = document.getElementById(`step_${idx}_mv`) || document.querySelector(`[name="step_${idx}_mv"]`);
                if (mvField) mvField.value = stepMv + " mV";
                
                let ldField = document.getElementById(`step_${idx}_load`) || document.querySelector(`[name="step_${idx}_load"]`);
                if (ldField) ldField.value = stepLoad + " t";
            });
        };
        
        // Removed periodic setInterval. Trigger only on user input/change:
        document.body.addEventListener('input', (e) => {
            if (e.target.id === 'final_mv_output' || e.target.name === 'final_mv_output' || e.target.name === 'capacity') {
                calcMvSteps();
            }
        });
    }
});

// Global handler to close PDF Preview Modal
document.addEventListener('DOMContentLoaded', () => {
    const attachModalCloseListeners = () => {
        const btnClose = document.getElementById('btnClosePreviewModal');
        const modal = document.getElementById('pdfPreviewModal');
        
        const closeModal = () => {
            if (modal) modal.style.display = 'none';
        };

        if (btnClose) {
            btnClose.addEventListener('click', closeModal);
        }

        if (modal) {
            modal.addEventListener('click', (e) => {
                if (e.target === modal || e.target.id === 'pdfPreviewModal') {
                    closeModal();
                }
            });
        }
    };

    // Attach immediately
    attachModalCloseListeners();
    // And also attach on body clicks in case modal is dynamically injected
    document.body.addEventListener('click', (e) => {
        if (e.target.closest('#btnClosePreviewModal')) {
            const m = document.getElementById('pdfPreviewModal');
            if (m) m.style.display = 'none';
        }
        if (e.target.id === 'pdfPreviewModal') {
            e.target.style.display = 'none';
        }
    });
});


// Universal Auto-fill when typing Job No and pressing Enter, or moving to next field (blur / change)
document.addEventListener('DOMContentLoaded', () => {
    const jobInputSelector = 'input[name="job_no"], input[name="job_number"], input[id="job_no"], input[id="job_number"]';
    const jobNoInputs = document.querySelectorAll(jobInputSelector);
    let jobFetchDebounceTimer = null;
    let isJobFetching = false;
    
    jobNoInputs.forEach(input => {
        let lastAutoFetched = '';
        const triggerJobFetch = () => {
            const val = input.value.trim();
            if (!val || val === lastAutoFetched || isJobFetching) return;
            
            clearTimeout(jobFetchDebounceTimer);
            jobFetchDebounceTimer = setTimeout(async () => {
                if (val === lastAutoFetched || isJobFetching) return;
                lastAutoFetched = val;
                isJobFetching = true;
                try {
                    if (typeof window.fetchJobNoData === 'function') {
                        await window.fetchJobNoData();
                    } else if (typeof window.fetchExcelData === 'function') {
                        await window.fetchExcelData();
                    }
                } finally {
                    isJobFetching = false;
                }
            }, 300);
        };

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                clearTimeout(jobFetchDebounceTimer);
                const val = input.value.trim();
                if (val && val !== lastAutoFetched && !isJobFetching) {
                    lastAutoFetched = val;
                    if (typeof window.fetchJobNoData === 'function') {
                        window.fetchJobNoData();
                    } else if (typeof window.fetchExcelData === 'function') {
                        window.fetchExcelData();
                    }
                }
            }
        });

        input.addEventListener('blur', triggerJobFetch);
        input.addEventListener('change', triggerJobFetch);
    });
});

// Universal Multi-System Suite (Sticky Command Bar, Step 1 Cards, Overview Modal)
window.renderMultiSystemSelector = function(multiSystems, onSelect, systemsMeta) {
    if (!multiSystems || typeof multiSystems !== 'object') {
        const oldBanner = document.getElementById('multiSystemStickyBanner');
        if (oldBanner) oldBanner.remove();
        const oldGrid = document.getElementById('multiSystemCardsContainer');
        if (oldGrid) oldGrid.remove();
        return;
    }

    const sysKeys = Object.keys(multiSystems);
    if (sysKeys.length <= 1) {
        const oldBanner = document.getElementById('multiSystemStickyBanner');
        if (oldBanner) oldBanner.remove();
        const oldGrid = document.getElementById('multiSystemCardsContainer');
        if (oldGrid) oldGrid.remove();
        return;
    }

    // Build synthesized metadata if not provided
    let metaList = systemsMeta;
    if (!Array.isArray(metaList) || metaList.length !== sysKeys.length) {
        metaList = sysKeys.map((k, idx) => {
            const s = multiSystems[k] || {};
            const cap = s.capacity || '';
            const tag = s.tag_no || s.conveyor_no || '';
            const lc_m = s.lc_model || s.sensor_model || '';
            const lc_q = s.lc_qty || '';
            const di_m = s.di_model || s.indicator_model || '';
            const jb_m = s.jbox_model || s.jbox_model1 || '';
            const parts = [];
            if (cap) parts.append ? parts.append(cap) : parts.push(cap);
            if (lc_m) parts.push(lc_q ? `${lc_q}× ${lc_m}` : lc_m);
            if (di_m) parts.push(di_m);
            return {
                id: `sys_${idx + 1}`,
                key: k,
                title: k,
                tag_no: tag,
                capacity: cap,
                lc_model: lc_m,
                lc_qty: lc_q,
                di_model: di_m,
                jbox_model: jb_m,
                summary: parts.join(' • ') || 'Complete System'
            };
        });
    }

    let activeIdx = 0;
    if (window.__activeSystemKey) {
        const found = sysKeys.indexOf(window.__activeSystemKey);
        if (found !== -1) activeIdx = found;
    }

    // ── 1. RENDER STICKY COMMAND BAR (Visible Across Steps 1 - 5) ──
    let banner = document.getElementById('multiSystemStickyBanner');
    if (!banner) {
        banner = document.createElement('div');
        banner.id = 'multiSystemStickyBanner';
        banner.className = 'ms-sticky-banner';

        // Insert before stepperBar or step-box-1 or form
        const stepper = document.getElementById('stepperBar');
        const step1 = document.getElementById('step-box-1') || document.querySelector('.section-box') || document.querySelector('form');
        if (stepper && stepper.parentNode) {
            stepper.parentNode.insertBefore(banner, stepper.nextSibling);
        } else if (step1 && step1.parentNode) {
            step1.parentNode.insertBefore(banner, step1);
        } else {
            document.body.prepend(banner);
        }
    }

    // ── 2. RENDER STEP 1 CARDS CONTAINER ──
    let cardsContainer = document.getElementById('multiSystemCardsContainer');
    const step1Box = document.getElementById('step-box-1') || document.getElementById('step-1');
    if (step1Box) {
        if (!cardsContainer) {
            cardsContainer = document.createElement('div');
            cardsContainer.id = 'multiSystemCardsContainer';
            cardsContainer.className = 'ms-cards-container';
            const titleEl = step1Box.querySelector('.section-box-title') || step1Box.firstElementChild;
            if (titleEl && titleEl.nextSibling) {
                step1Box.insertBefore(cardsContainer, titleEl.nextSibling);
            } else {
                step1Box.prepend(cardsContainer);
            }
        }
    }

    function updateView() {
        const curMeta = metaList[activeIdx] || {};
        const curKey = curMeta.key || sysKeys[activeIdx];
        const curData = multiSystems[curKey] || {};
        window.__activeSystemKey = curKey;
        window.__activeSystemIndex = activeIdx;

        // 1. Render Sticky Command Bar HTML
        banner.innerHTML = `
            <div class="ms-banner-top">
                <div class="ms-banner-left">
                    <span class="badge-ms-count">
                        <i class="fa-solid fa-layer-group"></i> ${sysKeys.length} Systems Detected
                    </span>
                    <span class="badge-ms-active">
                        <i class="fa-solid fa-circle-check" style="color: #10b981;"></i> Current: 
                        <strong id="multiSysActiveTitle">${escapeHtml(curMeta.title || curKey)}</strong>
                    </span>
                </div>
                <div class="ms-banner-controls">
                    <button type="button" class="btn-ms-nav" id="btnMultiSysPrev" ${activeIdx === 0 ? 'disabled' : ''} title="Previous System">
                        <i class="fa-solid fa-chevron-left"></i> Prev
                    </button>
                    <select class="select-ms-dropdown" id="multiSystemSelectDropdown">
                        ${metaList.map((m, idx) => `
                            <option value="${idx}" ${idx === activeIdx ? 'selected' : ''}>
                                ${idx + 1}. ${escapeHtml(m.title || m.key)} [${escapeHtml(m.capacity || 'N/A')}]
                            </option>
                        `).join('')}
                    </select>
                    <button type="button" class="btn-ms-nav" id="btnMultiSysNext" ${activeIdx === sysKeys.length - 1 ? 'disabled' : ''} title="Next System">
                        Next <i class="fa-solid fa-chevron-right"></i>
                    </button>
                    <button type="button" class="btn-ms-overview" id="btnOpenSysOverview">
                        <i class="fa-solid fa-table-list"></i> All Systems (${sysKeys.length})
                    </button>
                </div>
            </div>
            <div class="ms-banner-pills">
                <span class="ms-spec-pill"><i class="fa-solid fa-tag"></i> Tag: <strong>${escapeHtml(curMeta.tag_no || 'N/A')}</strong></span>
                <span class="ms-spec-pill"><i class="fa-solid fa-weight-hanging"></i> Cap: <strong>${escapeHtml(curMeta.capacity || 'N/A')}</strong></span>
                <span class="ms-spec-pill"><i class="fa-solid fa-microchip"></i> LC: <strong>${escapeHtml((curMeta.lc_qty ? curMeta.lc_qty + '× ' : '') + (curMeta.lc_model || 'N/A'))}</strong></span>
                <span class="ms-spec-pill"><i class="fa-solid fa-gauge-high"></i> DI: <strong>${escapeHtml(curMeta.di_model || 'N/A')}</strong></span>
                ${curMeta.jbox_model ? `<span class="ms-spec-pill"><i class="fa-solid fa-box"></i> JB: <strong>${escapeHtml(curMeta.jbox_model)}</strong></span>` : ''}
            </div>
        `;

        // 2. Render Step 1 Cards HTML
        if (cardsContainer) {
            cardsContainer.innerHTML = `
                <div class="ms-cards-header">
                    <span style="font-size: 14px; font-weight: 800; color: #1e293b;">
                        <i class="fa-solid fa-cubes-stacked" style="color: #0284c7; margin-right: 6px;"></i> 
                        Select System to Auto-Fill & Test:
                    </span>
                    <span style="font-size: 12px; color: #64748b;">
                        Showing ${sysKeys.length} isolated systems from ST-ERP / Drawing
                    </span>
                </div>
                <div class="ms-cards-grid">
                    ${metaList.map((m, idx) => `
                        <div class="ms-card ${idx === activeIdx ? 'active' : ''}" data-idx="${idx}">
                            <div class="ms-card-top">
                                <span class="ms-card-index">System ${idx + 1} of ${sysKeys.length}</span>
                                <span class="ms-card-status">${idx === activeIdx ? 'ACTIVE' : 'SELECT'}</span>
                            </div>
                            <div class="ms-card-title" title="${escapeHtml(m.title || m.key)}">
                                ${escapeHtml(m.title || m.key)}
                            </div>
                            <div class="ms-card-specs-row">
                                ${m.tag_no ? `<span class="ms-card-mini-spec"><strong>Tag:</strong> ${escapeHtml(m.tag_no)}</span>` : ''}
                                ${m.capacity ? `<span class="ms-card-mini-spec"><strong>Cap:</strong> ${escapeHtml(m.capacity)}</span>` : ''}
                                ${m.lc_model ? `<span class="ms-card-mini-spec"><strong>LC:</strong> ${escapeHtml((m.lc_qty ? m.lc_qty + '× ' : '') + m.lc_model)}</span>` : ''}
                                ${m.di_model ? `<span class="ms-card-mini-spec"><strong>DI:</strong> ${escapeHtml(m.di_model)}</span>` : ''}
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;

            // Bind Card Clicks
            cardsContainer.querySelectorAll('.ms-card').forEach(card => {
                card.addEventListener('click', () => {
                    const idx = parseInt(card.getAttribute('data-idx'), 10);
                    if (!isNaN(idx) && idx !== activeIdx) {
                        selectSystem(idx);
                    }
                });
            });
        }

        // Bind Command Bar Events
        const btnPrev = banner.querySelector('#btnMultiSysPrev');
        const btnNext = banner.querySelector('#btnMultiSysNext');
        const selectDrop = banner.querySelector('#multiSystemSelectDropdown');
        const btnOverview = banner.querySelector('#btnOpenSysOverview');

        if (btnPrev) btnPrev.addEventListener('click', () => { if (activeIdx > 0) selectSystem(activeIdx - 1); });
        if (btnNext) btnNext.addEventListener('click', () => { if (activeIdx < sysKeys.length - 1) selectSystem(activeIdx + 1); });
        if (selectDrop) selectDrop.addEventListener('change', (e) => {
            const idx = parseInt(e.target.value, 10);
            if (!isNaN(idx)) selectSystem(idx);
        });
        if (btnOverview) btnOverview.addEventListener('click', openOverviewModal);
    }

    function selectSystem(idx) {
        if (idx < 0 || idx >= sysKeys.length) return;
        activeIdx = idx;
        const curMeta = metaList[activeIdx] || {};
        const curKey = curMeta.key || sysKeys[activeIdx];
        const curData = multiSystems[curKey] || {};

        updateView();

        if (typeof onSelect === 'function') {
            onSelect(curData, curKey);
        }

        if (typeof showToast === 'function') {
            showToast(`✅ Loaded System ${activeIdx + 1}: ${curMeta.title || curKey}`, 'info');
        }
    }

    function openOverviewModal() {
        let modal = document.getElementById('multiSystemOverviewModal');
        if (modal) modal.remove();

        modal = document.createElement('div');
        modal.id = 'multiSystemOverviewModal';
        modal.className = 'ms-modal-backdrop';
        modal.innerHTML = `
            <div class="ms-modal-card">
                <div class="ms-modal-header">
                    <div class="ms-modal-title">
                        <i class="fa-solid fa-layer-group" style="color: #38bdf8;"></i> 
                        All Detected Systems Directory (${sysKeys.length} Systems)
                    </div>
                    <button type="button" class="ms-modal-close" id="btnCloseMsModal">&times;</button>
                </div>
                <div class="ms-modal-search-bar">
                    <i class="fa-solid fa-magnifying-glass" style="color: #64748b;"></i>
                    <input type="text" class="ms-modal-search-input" id="msModalSearchInput" placeholder="Quick search by system name, tag number, capacity, load cell model...">
                </div>
                <div class="ms-modal-body">
                    <table class="ms-modal-table" id="msModalTable">
                        <thead>
                            <tr>
                                <th style="width: 50px;">#</th>
                                <th>System / Application</th>
                                <th>Tag / Conveyor</th>
                                <th>Capacity</th>
                                <th>Load Cell</th>
                                <th>Indicator</th>
                                <th style="text-align: right; width: 130px;">Action</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${metaList.map((m, idx) => `
                                <tr class="${idx === activeIdx ? 'row-active' : ''}" data-idx="${idx}" data-search="${escapeHtml(`${m.title} ${m.tag_no} ${m.capacity} ${m.lc_model} ${m.di_model}`).toLowerCase()}">
                                    <td><strong>${idx + 1}</strong></td>
                                    <td><strong style="color: #0f172a;">${escapeHtml(m.title || m.key)}</strong></td>
                                    <td>${escapeHtml(m.tag_no || '-')}</td>
                                    <td><span style="font-weight: 700; color: #0284c7;">${escapeHtml(m.capacity || '-')}</span></td>
                                    <td>${escapeHtml((m.lc_qty ? m.lc_qty + '× ' : '') + (m.lc_model || '-'))}</td>
                                    <td>${escapeHtml(m.di_model || '-')}</td>
                                    <td style="text-align: right;">
                                        <button type="button" class="btn-ms-select-row ${idx === activeIdx ? 'btn-selected' : ''}" data-idx="${idx}">
                                            ${idx === activeIdx ? '<i class="fa-solid fa-check"></i> Active' : 'Load System'}
                                        </button>
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
        document.body.appendChild(modal);

        // Bind Search
        const searchInput = modal.querySelector('#msModalSearchInput');
        searchInput.focus();
        searchInput.addEventListener('input', (e) => {
            const q = e.target.value.toLowerCase().trim();
            modal.querySelectorAll('#msModalTable tbody tr').forEach(row => {
                const searchStr = row.getAttribute('data-search') || '';
                row.style.display = searchStr.includes(q) ? '' : 'none';
            });
        });

        // Bind Select buttons & Close
        const closeModal = () => modal.remove();
        modal.querySelector('#btnCloseMsModal').addEventListener('click', closeModal);
        modal.addEventListener('click', (e) => { if (e.target === modal) closeModal(); });

        modal.querySelectorAll('.btn-ms-select-row').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const idx = parseInt(btn.getAttribute('data-idx'), 10);
                if (!isNaN(idx)) {
                    closeModal();
                    selectSystem(idx);
                }
            });
        });

        modal.querySelectorAll('#msModalTable tbody tr').forEach(row => {
            row.addEventListener('click', () => {
                const idx = parseInt(row.getAttribute('data-idx'), 10);
                if (!isNaN(idx)) {
                    closeModal();
                    selectSystem(idx);
                }
            });
        });
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    // Initial render
    updateView();
};


// ── Universal Multi-System Auto-Initializer & Switch Handler ──
window.autoInitMultiSystemSelector = function(json) {
    if (!json || !json.has_multiple_systems || !json.multi_systems) return;
    if (typeof window.renderMultiSystemSelector !== 'function') return;

    window.renderMultiSystemSelector(json.multi_systems, (sysData, sysName) => {
        // 1. Try template-specific spec application functions
        if (typeof window.applyExtractedDigitalIndicatorSpecs === 'function') {
            window.applyExtractedDigitalIndicatorSpecs(sysData);
        } else if (typeof window.applyExtractedBeltScaleSpecs === 'function') {
            window.applyExtractedBeltScaleSpecs(sysData);
        } else if (typeof window.applyExtractedCraneScaleSpecs === 'function') {
            window.applyExtractedCraneScaleSpecs(sysData);
        } else {
            // 2. Universal fallback: populate every input matching a spec key
            for (const [k, v] of Object.entries(sysData)) {
                if (v !== undefined && v !== null && v !== '') {
                    if (typeof window.setFieldValue === 'function') {
                        window.setFieldValue(k, v);
                    } else {
                        const el = document.getElementById(k) || document.querySelector(`[name="${k}"]`);
                        if (el) {
                            el.value = v;
                            try { el.dispatchEvent(new Event('change')); } catch(e){}
                        }
                    }
                }
            }
        }

        // 3. Recalculate calibration if function exists
        if (typeof window.autoCalculateCalibrationMatrix === 'function') {
            window.autoCalculateCalibrationMatrix(sysData.capacity || sysData.spec_4);
        }
    }, json.systems_meta);
};

// Intercept fetch('/api/fetch-job-data...') to ensure ALL templates display the Multi-System Command Bar
(function() {
    if (window.__multiSysFetchIntercepted) return;
    window.__multiSysFetchIntercepted = true;
    const _origFetch = window.fetch;
    window.fetch = async function(...args) {
        const res = await _origFetch.apply(this, args);
        try {
            const url = (typeof args[0] === 'string') ? args[0] : (args[0] && args[0].url ? args[0].url : '');
            if (url && url.includes('/api/fetch-job-data')) {
                const clone = res.clone();
                clone.json().then(json => {
                    if (json && json.status === 'success' && json.has_multiple_systems) {
                        setTimeout(() => window.autoInitMultiSystemSelector(json), 120);
                    }
                }).catch(() => {});
            }
        } catch(e) {}
        return res;
    };
})();

// Fallback GA File Input binder for templates without an inline handleGaFileUpload
document.addEventListener('DOMContentLoaded', () => {
    const gaInput = document.getElementById('gaFileInput');
    if (gaInput && typeof window.handleGaFileUpload !== 'function' && !gaInput.dataset.gaBound) {
        gaInput.dataset.gaBound = '1';
        gaInput.addEventListener('change', async (e) => {
            const file = e.target.files && e.target.files[0];
            if (!file) return;
            const formData = new FormData();
            formData.append('ga_file', file);
            try {
                const res = await fetch('/api/extract-ga', { method: 'POST', body: formData });
                const json = await res.json();
                if (json.status === 'success' && json.extracted) {
                    const ext = json.extracted;
                    const setter = (k, v) => {
                        if (v === undefined || v === null || v === '') return;
                        if (typeof window.setFieldValue === 'function') window.setFieldValue(k, v);
                        else {
                            const el = document.getElementById(k) || document.querySelector(`[name="${k}"]`);
                            if (el) { el.value = v; try { el.dispatchEvent(new Event('change')); } catch (_) {} }
                        }
                    };
                    for (const [k, v] of Object.entries(ext)) setter(k, v);
                    if (ext.bs_model) setter('indicator_model', ext.bs_model);
                    if (ext.jbox_model1) setter('jbox_model', ext.jbox_model1);
                    if (ext.remote_model) setter('remote_display_model', ext.remote_model);
                    if (typeof window.autoCalculateCalibrationMatrix === 'function' && ext.capacity) {
                        window.autoCalculateCalibrationMatrix(ext.capacity, ext.belt_speed);
                    }
                    if (typeof showToast === 'function') showToast(`Auto-filled from ${json.filename}`, 'success');
                } else if (typeof showToast === 'function') {
                    showToast(json.message || 'Could not extract GA file.', 'error');
                }
            } catch (err) {
                if (typeof showToast === 'function') showToast('GA extract failed: ' + err.message, 'error');
            } finally {
                e.target.value = '';
            }
        });
    }
});
