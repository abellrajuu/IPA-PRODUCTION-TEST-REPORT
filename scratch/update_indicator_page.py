import os

file_path = "frontend_html_templates/digital_indicator.html"
with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

old_func = """        function setFieldValue(id, val) {
            if (val === undefined || val === null || val === '') return;
            const el = document.getElementById(id) || document.getElementsByName(id)[0];
            if (el) {
                el.value = val;
                el.style.transition = 'all 0.3s ease';
                el.style.border = '2px solid #10b981';
                el.style.background = '#ecfdf5';
                setTimeout(() => {
                    el.style.border = '';
                    el.style.background = '';
                }, 3500);
            }
        }

        // 5-Step Load & mV Calibration Matrix Calculator
        function autoCalculateCalibrationMatrix(capVal) {
            const modeInput = document.getElementById('calib_mode');
            const isLoad = modeInput && modeInput.value === 'load';
            const capInput = document.getElementById('capacity');
            const rawCap = capVal || (capInput ? capInput.value.trim() : '');
            
            let spanVal = 10.0;
            let unit = isLoad ? 'Kg' : 'mV';
            
            if (rawCap) {
                const num = parseFloat(String(rawCap).replace(/[^0-9.]/g, ''));
                if (!isNaN(num) && num > 0) {
                    if (isLoad) {
                        spanVal = num;
                        const up = rawCap.toUpperCase();
                        if (up.includes('TON') || up.includes(' T') || up.endsWith('T')) unit = 'T';
                        else if (up.includes('KG')) unit = 'Kg';
                        else unit = 'Kg';
                    } else {
                        // In mV mode, if rawCap explicitly has mV, use it; otherwise standard 10.00 mV span
                        if (rawCap.toLowerCase().includes('mv') && num <= 50) {
                            spanVal = num;
                        } else {
                            spanVal = 10.0;
                        }
                        unit = 'mV';
                    }
                }
            } else {
                spanVal = isLoad ? 1000.0 : 10.0;
                unit = isLoad ? 'Kg' : 'mV';
            }

            const dec = (unit === 'mV' || spanVal < 100) ? 2 : (spanVal < 1000 ? 1 : 0);

            // 1. Populate Routine Test Table Row 4 & Row 5 (Specified and Checked)
            setFieldValue('spec_4', `${spanVal.toFixed(dec)} ${unit}`);
            setFieldValue('act_4', `${spanVal.toFixed(dec)} ${unit}`);
            setFieldValue('spec_5', 'NA');
            setFieldValue('act_5', 'NA');

            // 2. Compute 5-Step Readings (20%, 40%, 60%, 80%, 100% of Span)
            // User readings on the right side are the step values (Value divided by 5 into 5 equal steps)
            const stepRatio = spanVal / 5.0;
            const pcts = [0.20, 0.40, 0.60, 0.80, 1.00];

            pcts.forEach((pct, idx) => {
                const r = idx + 1; // Rows 1 to 5
                const stepNumeric = (stepRatio * (idx + 1)).toFixed(dec);
                const stepDisplay = `${stepNumeric} ${unit}`;
                
                // Left column: Applied Step Load / mV
                setFieldValue(`g${r}_0`, stepDisplay);

                // Load Cells 1 through 4 (columns g{r}_1 to g{r}_4)
                for (let lc = 1; lc <= 4; lc++) {
                    setFieldValue(`g${r}_${lc}`, stepNumeric);
                }

                // Current output mA: 4.00 mA to 20.00 mA (column g{r}_5)
                const maVal = (4.0 + pct * 16.0).toFixed(2);
                setFieldValue(`g${r}_5`, maVal);

                // Load Cells 5 through 8 (columns g{r}_6 to g{r}_9)
                for (let lc = 6; lc <= 9; lc++) {
                    setFieldValue(`g${r}_${lc}`, stepNumeric);
                }
            });

            // 3. Clear extra rows 6 to 12 if not customized
            for (let r = 6; r <= 12; r++) {
                const g0 = document.getElementById(`g${r}_0`);
                if (g0 && (!g0.value || g0.value.includes('%'))) {
                    g0.value = '';
                }
            }

            // 4. Populate Span / Tare / Excitation table
            for (let lc = 1; lc <= 8; lc++) {
                setFieldValue(`span_lc${lc}`, spanVal.toFixed(dec));
                setFieldValue(`tare_lc${lc}`, '0.00');
                setFieldValue(`excv_lc${lc}`, '10.0V');
            }
        }"""

new_func = """        function setFieldValue(id, val) {
            if (val === undefined || val === null || val === '') return;
            const el = document.getElementById(id) || document.getElementsByName(id)[0];
            if (el) {
                // If this is the active input element being typed in by the user, do not overwrite cursor
                if (document.activeElement === el) return;
                el.value = val;
                el.style.transition = 'all 0.3s ease';
                el.style.border = '2px solid #10b981';
                el.style.background = '#ecfdf5';
                setTimeout(() => {
                    el.style.border = '';
                    el.style.background = '';
                }, 3500);
            }
        }

        // 5-Step Load & mV Calibration Matrix Calculator
        function autoCalculateCalibrationMatrix(sourceVal, isExplicitRow4) {
            const modeInput = document.getElementById('calib_mode');
            const isLoad = modeInput && modeInput.value === 'load';
            const capInput = document.getElementById('capacity');
            const spec4Input = document.getElementById('spec_4');
            const act4Input = document.getElementById('act_4');

            // Determine raw input string to use for span
            let rawStr = '';
            if (isExplicitRow4 && sourceVal !== undefined && sourceVal !== null) {
                rawStr = String(sourceVal).trim();
            } else if (sourceVal !== undefined && sourceVal !== null && String(sourceVal).trim()) {
                rawStr = String(sourceVal).trim();
            } else if (spec4Input && spec4Input.value && spec4Input.value.trim() && !spec4Input.value.toLowerCase().includes('sample') && !spec4Input.value.toLowerCase().includes('span')) {
                rawStr = spec4Input.value.trim();
            } else if (act4Input && act4Input.value && act4Input.value.trim() && !act4Input.value.toLowerCase().includes('passed')) {
                rawStr = act4Input.value.trim();
            } else if (capInput && capInput.value.trim()) {
                rawStr = capInput.value.trim();
            }

            let spanVal = isLoad ? 1000.0 : 10.0;
            let unit = isLoad ? 'Kg' : 'mV';

            if (rawStr) {
                const num = parseFloat(rawStr.replace(/[^0-9.]/g, ''));
                if (!isNaN(num) && num > 0) {
                    spanVal = num;
                    const up = rawStr.toUpperCase();
                    if (up.includes('TON') || up.includes(' T') || up.endsWith('T')) unit = 'T';
                    else if (up.includes('KG')) unit = 'Kg';
                    else if (up.includes('G') && !up.includes('KG')) unit = 'g';
                    else if (up.includes('MV')) unit = 'mV';
                    else {
                        unit = isLoad ? 'Kg' : 'mV';
                    }
                }
            }

            // Auto-switch mode toggle if unit clearly indicates mode
            if (unit === 'mV' && isLoad) {
                if (modeInput) modeInput.value = 'mv';
                const btnMv = document.getElementById('btnModeMv');
                const btnLoad = document.getElementById('btnModeLoad');
                if (btnMv) btnMv.classList.add('active');
                if (btnLoad) btnLoad.classList.remove('active');
                const spanBox = document.getElementById('spanTareExcvContainer');
                if (spanBox) spanBox.style.display = 'none';
            } else if ((unit === 'Kg' || unit === 'T' || unit === 'g') && !isLoad) {
                if (modeInput) modeInput.value = 'load';
                const btnMv = document.getElementById('btnModeMv');
                const btnLoad = document.getElementById('btnModeLoad');
                if (btnLoad) btnLoad.classList.add('active');
                if (btnMv) btnMv.classList.remove('active');
                const spanBox = document.getElementById('spanTareExcvContainer');
                if (spanBox) spanBox.style.display = 'block';
            }

            const dec = (unit === 'mV' || spanVal < 100) ? 2 : (spanVal < 1000 ? 1 : 0);
            const formattedSpan = `${spanVal.toFixed(dec)} ${unit}`;

            // 1. Synchronize Routine Test Table Row 4 & Row 5
            if (isExplicitRow4) {
                if (document.activeElement === spec4Input) {
                    if (act4Input) act4Input.value = spec4Input.value;
                } else if (document.activeElement === act4Input) {
                    if (spec4Input) spec4Input.value = act4Input.value;
                } else {
                    if (spec4Input) spec4Input.value = formattedSpan;
                    if (act4Input) act4Input.value = formattedSpan;
                }
            } else {
                setFieldValue('spec_4', formattedSpan);
                setFieldValue('act_4', formattedSpan);
            }
            setFieldValue('spec_5', 'NA');
            setFieldValue('act_5', 'NA');

            // 2. Compute 5-Step Readings (20%, 40%, 60%, 80%, 100% of Span)
            // Left column is stepDisplay, Load cells are stepNumeric (divided by 5 equal steps)
            const stepRatio = spanVal / 5.0;
            const pcts = [0.20, 0.40, 0.60, 0.80, 1.00];

            pcts.forEach((pct, idx) => {
                const r = idx + 1; // Rows 1 to 5
                const stepNumeric = (stepRatio * (idx + 1)).toFixed(dec);
                const stepDisplay = `${stepNumeric} ${unit}`;

                // Left column: Applied Step Load / mV
                setFieldValue(`g${r}_0`, stepDisplay);

                // Load Cells 1 through 4 (columns g{r}_1 to g{r}_4)
                for (let lc = 1; lc <= 4; lc++) {
                    setFieldValue(`g${r}_${lc}`, stepNumeric);
                }

                // Current output mA: 4.00 mA to 20.00 mA (column g{r}_5)
                const maVal = (4.0 + pct * 16.0).toFixed(2);
                setFieldValue(`g${r}_5`, maVal);

                // Load Cells 5 through 8 (columns g{r}_6 to g{r}_9)
                for (let lc = 6; lc <= 9; lc++) {
                    setFieldValue(`g${r}_${lc}`, stepNumeric);
                }
            });

            // 3. Clear extra rows 6 to 12 if not customized
            for (let r = 6; r <= 12; r++) {
                const g0 = document.getElementById(`g${r}_0`);
                if (g0 && (!g0.value || g0.value.includes('%'))) {
                    g0.value = '';
                }
            }

            // 4. Populate Span / Tare / Excitation table
            for (let lc = 1; lc <= 8; lc++) {
                setFieldValue(`span_lc${lc}`, spanVal.toFixed(dec));
                setFieldValue(`tare_lc${lc}`, '0.00');
                setFieldValue(`excv_lc${lc}`, '10.0V');
            }
        }"""

# Normalize CRLF
content_norm = content.replace("\r\n", "\n")
old_func_norm = old_func.replace("\r\n", "\n")
new_func_norm = new_func.replace("\r\n", "\n")

if old_func_norm in content_norm:
    content_norm = content_norm.replace(old_func_norm, new_func_norm)
    print("Function replaced successfully!")
else:
    print("ERROR: old_func_norm not found in content!")

# Now add listeners for spec_4 and act_4
old_listeners = """            const capInput = document.getElementById('capacity');
            if (capInput) {
                capInput.addEventListener('input', () => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix();
                });
                capInput.addEventListener('change', () => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix();
                });
            }"""

new_listeners = """            const capInput = document.getElementById('capacity');
            if (capInput) {
                capInput.addEventListener('input', () => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix(capInput.value, false);
                });
                capInput.addEventListener('change', () => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix(capInput.value, false);
                });
            }

            const spec4Input = document.getElementById('spec_4');
            const act4Input = document.getElementById('act_4');
            if (spec4Input) {
                spec4Input.addEventListener('input', (e) => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix(e.target.value, true);
                });
                spec4Input.addEventListener('change', (e) => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix(e.target.value, true);
                });
            }
            if (act4Input) {
                act4Input.addEventListener('input', (e) => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix(e.target.value, true);
                });
                act4Input.addEventListener('change', (e) => {
                    if (typeof autoCalculateCalibrationMatrix === 'function') autoCalculateCalibrationMatrix(e.target.value, true);
                });
            }"""

old_listeners_norm = old_listeners.replace("\r\n", "\n")
new_listeners_norm = new_listeners.replace("\r\n", "\n")

if old_listeners_norm in content_norm:
    content_norm = content_norm.replace(old_listeners_norm, new_listeners_norm)
    print("Listeners added successfully!")
else:
    print("ERROR: old_listeners_norm not found in content!")

# Also update applyExtractedDigitalIndicatorSpecs to pass spec_4 || act_4 || capacity
old_apply = """            if (typeof autoCalculateCalibrationMatrix === 'function') {
                autoCalculateCalibrationMatrix(data.capacity);
            }"""
new_apply = """            if (typeof autoCalculateCalibrationMatrix === 'function') {
                autoCalculateCalibrationMatrix(data.spec_4 || data.act_4 || data.capacity, true);
            }"""

old_apply_norm = old_apply.replace("\r\n", "\n")
new_apply_norm = new_apply.replace("\r\n", "\n")

if old_apply_norm in content_norm:
    content_norm = content_norm.replace(old_apply_norm, new_apply_norm)
    print("applyExtracted updated successfully!")
else:
    print("ERROR: old_apply_norm not found in content!")

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content_norm)

print("Saved updated digital_indicator.html!")
