import sys, os

portal_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\portal.html'
with open(portal_path, 'r', encoding='utf-8') as f:
    html = f.read()

# Let's inspect CSS insertion point
style_end = html.find('</style>')
if style_end == -1:
    print('Error: </style> not found')
    sys.exit(1)

bento_css = """
        /* ══════════════════════════════════════════════════════════════════
           UI-UX PRO MAX: BENTO BOX WORKSTATION & OMNI-SEARCH STYLES
        ══════════════════════════════════════════════════════════════════ */
        .telemetry-strip {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 12px;
            flex-wrap: wrap;
            margin-bottom: 22px;
        }
        .telemetry-chip {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 6px 14px;
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 999px;
            font-size: 12px;
            font-weight: 700;
            color: #334155;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        }
        .telemetry-chip i { font-size: 11px; }
        .telemetry-chip.chip-online {
            background: #f0fdf4;
            border-color: #bbf7d0;
            color: #166534;
        }
        .telemetry-chip.chip-online i { color: #16a34a; }
        .telemetry-chip.chip-tech {
            background: #f0f9ff;
            border-color: #bae6fd;
            color: #0369a1;
        }
        .telemetry-chip.chip-tech i { color: #0284c7; }

        /* Omni Search Box */
        .omni-search-wrapper {
            max-width: 720px;
            margin: 0 auto 28px;
            position: relative;
        }
        .omni-search-input-wrap {
            display: flex;
            align-items: center;
            background: #ffffff;
            border: 2px solid #e2e8f0;
            border-radius: 16px;
            padding: 6px 8px 6px 18px;
            box-shadow: 0 8px 24px -4px rgba(15, 23, 42, 0.06), 0 2px 6px -1px rgba(15, 23, 42, 0.04);
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }
        .omni-search-input-wrap:focus-within {
            border-color: #0284c7;
            box-shadow: 0 12px 32px -4px rgba(2, 132, 199, 0.2), 0 0 0 4px rgba(2, 132, 199, 0.12);
            transform: translateY(-2px);
        }
        .omni-search-input-wrap i.search-main-icon {
            font-size: 18px;
            color: #0284c7;
            margin-right: 14px;
        }
        .omni-search-input {
            flex: 1;
            border: none;
            outline: none;
            font-family: 'Inter', sans-serif;
            font-size: 15px;
            font-weight: 600;
            color: #0f172a;
            background: transparent;
        }
        .omni-search-input::placeholder {
            color: #94a3b8;
            font-weight: 500;
        }
        .omni-search-btn {
            background: linear-gradient(135deg, #0284c7, #0369a1);
            color: #ffffff;
            border: none;
            padding: 10px 22px;
            border-radius: 12px;
            font-family: 'Inter', sans-serif;
            font-size: 13.5px;
            font-weight: 700;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 4px 12px rgba(2, 132, 199, 0.35);
            transition: all 0.2s ease;
        }
        .omni-search-btn:hover {
            transform: scale(1.02);
            box-shadow: 0 6px 16px rgba(2, 132, 199, 0.45);
        }

        /* Omni Suggestions Dropdown */
        .omni-results-dropdown {
            position: absolute;
            top: calc(100% + 8px);
            left: 0;
            right: 0;
            background: #ffffff;
            border: 1.5px solid #e2e8f0;
            border-radius: 16px;
            box-shadow: 0 16px 36px -6px rgba(15, 23, 42, 0.16);
            z-index: 1000;
            max-height: 380px;
            overflow-y: auto;
            display: none;
            padding: 8px;
            animation: cardFadeIn 0.2s ease-out;
        }
        .omni-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 12px 16px;
            border-radius: 10px;
            cursor: pointer;
            text-decoration: none;
            color: #0f172a;
            transition: all 0.15s ease;
        }
        .omni-item:hover {
            background: #f0f9ff;
            color: #0284c7;
        }
        .omni-item-left {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .omni-item-icon {
            width: 36px;
            height: 36px;
            border-radius: 8px;
            background: #e0f2fe;
            color: #0284c7;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 15px;
        }
        .omni-item-title {
            font-weight: 700;
            font-size: 14px;
        }
        .omni-item-sub {
            font-size: 12px;
            color: #64748b;
        }
        .omni-item-badge {
            font-size: 11px;
            font-weight: 700;
            padding: 3px 8px;
            border-radius: 6px;
            background: #f1f5f9;
            color: #475569;
        }

        /* Quick-Launch Strip */
        .quick-strip-label {
            text-align: center;
            font-size: 11.5px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: #64748b;
            margin-bottom: 10px;
        }
        .quick-launch-strip {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            flex-wrap: wrap;
            margin-bottom: 34px;
        }
        .quick-chip {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 9px 18px;
            background: #ffffff;
            border: 1.5px solid #e2e8f0;
            border-radius: 12px;
            color: #1e293b;
            font-size: 13.5px;
            font-weight: 700;
            text-decoration: none;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
            transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
            cursor: pointer;
        }
        .quick-chip i {
            color: #0284c7;
            font-size: 14px;
        }
        .quick-chip:hover {
            border-color: #0284c7;
            background: #f0f9ff;
            color: #0284c7;
            transform: translateY(-2px);
            box-shadow: 0 6px 16px -2px rgba(2, 132, 199, 0.18);
        }

        /* Bento Grid 2026 */
        .bento-workstation-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(300px, 1fr));
            gap: 22px;
            max-width: 980px;
            margin: 0 auto;
        }
        @media (max-width: 768px) {
            .bento-workstation-grid {
                grid-template-columns: 1fr;
            }
        }
        .bento-card {
            background: #ffffff;
            border: 1.5px solid #e2e8f0;
            border-radius: 20px;
            padding: 28px 26px;
            cursor: pointer;
            text-decoration: none;
            color: inherit;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 6px 20px -4px rgba(15, 23, 42, 0.05);
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
            position: relative;
            overflow: hidden;
        }
        .bento-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 4px;
            background: transparent;
            transition: background 0.25s ease;
        }
        .bento-card:hover {
            border-color: #0284c7;
            transform: translateY(-4px);
            box-shadow: 0 16px 36px -6px rgba(2, 132, 199, 0.16);
        }
        .bento-card:hover::before {
            background: linear-gradient(90deg, #0284c7, #38bdf8);
        }
        .bento-card-top {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            margin-bottom: 18px;
        }
        .bento-icon {
            width: 58px;
            height: 58px;
            border-radius: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 26px;
            transition: transform 0.25s ease;
        }
        .bento-card:hover .bento-icon {
            transform: scale(1.08);
        }
        .bento-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 11.5px;
            font-weight: 800;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            padding: 5px 12px;
            border-radius: 999px;
        }
        .bento-title {
            font-size: 20px;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 8px;
            letter-spacing: -0.3px;
        }
        .bento-desc {
            font-size: 13.5px;
            color: #64748b;
            line-height: 1.5;
            margin-bottom: 22px;
        }
        .bento-footer {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-top: 16px;
            border-top: 1px solid #f1f5f9;
        }
        .bento-action-link {
            font-size: 13.5px;
            font-weight: 800;
            color: #0284c7;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .bento-card:hover .bento-action-link i {
            transform: translateX(4px);
            transition: transform 0.2s ease;
        }

        /* Specific Bento Color Accents */
        .bento-primary .bento-icon {
            background: linear-gradient(135deg, #e0f2fe, #bae6fd);
            color: #0284c7;
            border: 1.5px solid #bae6fd;
        }
        .bento-primary .bento-badge {
            background: #e0f2fe;
            color: #0284c7;
            border: 1px solid #bae6fd;
        }
        .bento-amber .bento-icon {
            background: linear-gradient(135deg, #fef3c7, #fde68a);
            color: #d97706;
            border: 1.5px solid #fde68a;
        }
        .bento-amber .bento-badge {
            background: #fef3c7;
            color: #d97706;
            border: 1px solid #fde68a;
        }
        .bento-archive .bento-icon {
            background: linear-gradient(135deg, #f1f5f9, #e2e8f0);
            color: #334155;
            border: 1.5px solid #cbd5e1;
        }
        .bento-archive .bento-badge {
            background: #f1f5f9;
            color: #475569;
            border: 1px solid #cbd5e1;
        }
        .bento-mobile .bento-icon {
            background: linear-gradient(135deg, #f0fdf4, #dcfce7);
            color: #16a34a;
            border: 1.5px solid #bbf7d0;
        }
        .bento-mobile .bento-badge {
            background: #f0fdf4;
            color: #166534;
            border: 1px solid #bbf7d0;
        }
"""

html_with_css = html[:style_end] + bento_css + html[style_end:]

# Now replace view-home
vh_start = html_with_css.find('<div id="view-home"')
vh_end = html_with_css.find('<div id="view-test-reports-hub"')
if vh_start == -1 or vh_end == -1:
    print('Error: view-home or view-test-reports-hub not found')
    sys.exit(1)

new_view_home = """<div id="view-home" class="view active">
        <!-- WORKSTATION BRAND HEADER -->
        <div class="workstation-header" style="margin-bottom: 18px;">
            <h1 class="workstation-title" style="font-size: 32px; font-weight: 900; color: #0c2340; letter-spacing: -0.8px;">Production Test Workstation</h1>
            <p class="workstation-sub" style="font-size: 15px; color: #64748b; font-weight: 500;">IPA Weighing &amp; Automation • Factory Calibration &amp; Test Reporting</p>
        </div>

        <!-- TELEMETRY & SYSTEM HEALTH STRIP -->
        <div class="telemetry-strip">
            <span class="telemetry-chip chip-online">
                <i class="fa-solid fa-circle-check"></i> ST-ERP SQL Server Online
            </span>
            <span class="telemetry-chip chip-tech">
                <i class="fa-solid fa-layer-group"></i> 15 Certified Systems Active
            </span>
            <span class="telemetry-chip">
                <i class="fa-solid fa-bolt"></i> Calibration Matrix Ready
            </span>
            <span class="telemetry-chip">
                <i class="fa-solid fa-calendar-day"></i> <span id="portalDateDisp"></span>
            </span>
        </div>

        <!-- OMNI-SEARCH & QUICK-LAUNCHER -->
        <div class="omni-search-wrapper">
            <div class="omni-search-input-wrap">
                <i class="fa-solid fa-magnifying-glass search-main-icon"></i>
                <input type="text" id="omniSearchInput" class="omni-search-input" placeholder="Search by Job No (e.g. SY0192, SY1246), Customer, or Equipment..." autocomplete="off">
                <button type="button" id="omniSearchBtn" class="omni-search-btn" onclick="executeOmniSearch()">
                    <span>Launch</span> <i class="fa-solid fa-arrow-right"></i>
                </button>
            </div>
            <!-- Live Suggestions Dropdown -->
            <div id="omniResultsDropdown" class="omni-results-dropdown"></div>
        </div>

        <!-- QUICK ACCESS PINNED EQUIPMENT STRIP -->
        <div class="quick-strip-label">Frequently Used Equipment</div>
        <div class="quick-launch-strip">
            <a href="/digital-indicator" class="quick-chip">
                <i class="fa-solid fa-microchip"></i> Digital Indicator
            </a>
            <a href="/belt-scale" class="quick-chip">
                <i class="fa-solid fa-truck-ramp-box"></i> Belt Scale
            </a>
            <a href="/crane-scale" class="quick-chip">
                <i class="fa-solid fa-weight-hanging"></i> Crane Scale
            </a>
            <a href="/weigh-feeder" class="quick-chip">
                <i class="fa-solid fa-boxes-stacked"></i> Weigh Feeder
            </a>
            <a href="/signal-conditioner" class="quick-chip">
                <i class="fa-solid fa-wave-square"></i> Signal Conditioner
            </a>
            <a href="/misc-report" class="quick-chip">
                <i class="fa-solid fa-file-lines"></i> Misc Report
            </a>
        </div>

        <!-- BENTO BOX WORKSTATION CARDS -->
        <div class="bento-workstation-grid">

            <!-- BENTO 1: PRODUCTION SYSTEMS HUB -->
            <div class="bento-card bento-primary" onclick="openView('systems-production')">
                <div class="bento-card-top">
                    <div class="bento-icon">
                        <i class="fa-solid fa-industry"></i>
                    </div>
                    <span class="bento-badge">
                        <i class="fa-solid fa-layer-group"></i> 15 Systems
                    </span>
                </div>
                <div>
                    <h3 class="bento-title">Production Systems</h3>
                    <p class="bento-desc">Complete test reports for Belt Scales, Digital Indicators, Crane Scales, Weigh Feeders, Loss-in-Weight, Batching, and Vibration instruments.</p>
                </div>
                <div class="bento-footer">
                    <span class="bento-action-link">
                        Explore Systems <i class="fa-solid fa-arrow-right"></i>
                    </span>
                    <span style="font-size: 12px; color: #94a3b8; font-weight: 600;">2-Page Word &amp; PDF</span>
                </div>
            </div>

            <!-- BENTO 2: ALSTOM TRACTION CONTROLLERS -->
            <div class="bento-card bento-amber" onclick="openView('systems-alstom')">
                <div class="bento-card-top">
                    <div class="bento-icon">
                        <i class="fa-solid fa-microchip"></i>
                    </div>
                    <span class="bento-badge">
                        <i class="fa-solid fa-cpu"></i> 2 Controllers
                    </span>
                </div>
                <div>
                    <h3 class="bento-title">Alstom Controllers</h3>
                    <p class="bento-desc">Dedicated validation protocols for ODD System (PP-05_18_A) and DD System (PP-05_18_B) locomotive traction weight transmitters.</p>
                </div>
                <div class="bento-footer">
                    <span class="bento-action-link" style="color: #d97706;">
                        Open Controllers <i class="fa-solid fa-arrow-right"></i>
                    </span>
                    <span style="font-size: 12px; color: #94a3b8; font-weight: 600;">Traction Specs</span>
                </div>
            </div>

            <!-- BENTO 3: SAVED TEST REPORTS & APPROVALS -->
            <div class="bento-card bento-archive" onclick="window.location.href='/review-dashboard'">
                <div class="bento-card-top">
                    <div class="bento-icon">
                        <i class="fa-solid fa-database"></i>
                    </div>
                    <span class="bento-badge">
                        <i class="fa-solid fa-signature"></i> Verifications
                    </span>
                </div>
                <div>
                    <h3 class="bento-title">Saved Reports &amp; Archive</h3>
                    <p class="bento-desc">View all saved production test records, verify measurements, download official DOCX/PDF reports, and apply HOD e-signatures.</p>
                </div>
                <div class="bento-footer">
                    <span class="bento-action-link" style="color: #334155;">
                        View Test History <i class="fa-solid fa-arrow-right"></i>
                    </span>
                    <span style="font-size: 12px; color: #94a3b8; font-weight: 600;">Search &amp; Export</span>
                </div>
            </div>

            <!-- BENTO 4: FIELD MOBILITY & CONNECTIVITY -->
            <div class="bento-card bento-mobile" onclick="window.location.href='/download-app'">
                <div class="bento-card-top">
                    <div class="bento-icon">
                        <i class="fa-brands fa-android"></i>
                    </div>
                    <span class="bento-badge">
                        <i class="fa-solid fa-wifi"></i> Offline Ready
                    </span>
                </div>
                <div>
                    <h3 class="bento-title">Android App &amp; Offline Sync</h3>
                    <p class="bento-desc">Install the Production Report APK for 1-tap mobile operation on the shop floor with local SQLite caching and offline fallback.</p>
                </div>
                <div class="bento-footer">
                    <span class="bento-action-link" style="color: #16a34a;">
                        Download App (.APK) <i class="fa-solid fa-download"></i>
                    </span>
                    <span style="font-size: 12px; color: #94a3b8; font-weight: 600;">V3.1.3 Stable</span>
                </div>
            </div>

        </div>
    </div>

    """

final_html = html_with_css[:vh_start] + new_view_home + html_with_css[vh_end:]

# Now add Omni-Search JS right before </body>
omni_script = """
    <!-- UI-UX PRO MAX: OMNI SEARCH & LIVE DATE SCRIPT -->
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            const dateEl = document.getElementById('portalDateDisp');
            if (dateEl) {
                const now = new Date();
                dateEl.textContent = now.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
            }

            const searchInput = document.getElementById('omniSearchInput');
            const dropdown = document.getElementById('omniResultsDropdown');

            const ALL_SYSTEMS = [
                { name: 'Digital Indicator', url: '/digital-indicator', icon: 'fa-microchip', cat: 'Production', tags: 'di ds105 ds105a4 indicator' },
                { name: 'Belt Scale', url: '/belt-scale', icon: 'fa-truck-ramp-box', cat: 'Production', tags: 'belt scale bs conveyor weighing' },
                { name: 'Crane Scale', url: '/crane-scale', icon: 'fa-weight-hanging', cat: 'Production', tags: 'crane cws wireless shackle' },
                { name: 'Weigh Feeder', url: '/weigh-feeder', icon: 'fa-boxes-stacked', cat: 'Production', tags: 'weigh feeder wf belt feeder' },
                { name: 'Loss in Weigh Feeder', url: '/loss-in-weigh-feeder', icon: 'fa-filter', cat: 'Production', tags: 'loss in weigh liw feeder' },
                { name: 'Batching System', url: '/batching-system', icon: 'fa-cubes', cat: 'Production', tags: 'batching bch automatic batch' },
                { name: 'Signal Conditioner', url: '/signal-conditioner', icon: 'fa-wave-square', cat: 'Sensors', tags: 'signal conditioner sc load cell transmitter' },
                { name: 'Remote Indicator', url: '/remote-indicator', icon: 'fa-desktop', cat: 'Production', tags: 'remote indicator display rd' },
                { name: 'Miscellaneous Report', url: '/misc-report', icon: 'fa-file-lines', cat: 'Production', tags: 'misc load cell junction box remote' },
                { name: 'ODD System', url: '/odd-system', icon: 'fa-train', cat: 'Alstom', tags: 'odd alstom traction transmitter' },
                { name: 'DD System', url: '/dd-system', icon: 'fa-train-subway', cat: 'Alstom', tags: 'dd alstom dual transmitter' },
                { name: 'Trip Safe', url: '/trip-safe', icon: 'fa-shield-halved', cat: 'Safety', tags: 'trip safe safety relay' },
                { name: 'Vibration Switch', url: '/vibration-switch', icon: 'fa-bolt', cat: 'Sensors', tags: 'vibration switch sensor' },
                { name: 'Vibration Meter', url: '/vibration-meter', icon: 'fa-gauge-high', cat: 'Sensors', tags: 'vibration meter' },
                { name: 'ACC mV', url: '/acc-mv', icon: 'fa-wave-square', cat: 'Sensors', tags: 'accelerometer acc mv' },
                { name: 'ACC Charge', url: '/acc-charge', icon: 'fa-battery-three-quarters', cat: 'Sensors', tags: 'accelerometer charge' },
                { name: 'Charge Amplifier', url: '/charge-amplifier', icon: 'fa-microchip', cat: 'Sensors', tags: 'charge amplifier' },
                { name: 'Job Traveller Card', url: '/job-traveller-card', icon: 'fa-id-card', cat: 'Docs', tags: 'jtc job traveller routing' }
            ];

            if (searchInput && dropdown) {
                searchInput.addEventListener('input', function() {
                    const q = this.value.trim().toLowerCase();
                    if (!q) {
                        dropdown.style.display = 'none';
                        dropdown.innerHTML = '';
                        return;
                    }

                    // Check if query looks like a Job Number (e.g. SY0192, SY1246, JOB-001)
                    const isJobNo = /^(sy|job|bs|di|cs|wf|ord|wo)/i.test(q) || /\\d{3,}/.test(q);

                    let html = '';
                    if (isJobNo) {
                        html += `
                            <div style="padding: 8px 12px 6px; font-size: 11px; font-weight: 800; color: #0284c7; text-transform: uppercase; letter-spacing: 0.5px;">
                                <i class="fa-solid fa-bolt"></i> Fast Job Launch (${q.toUpperCase()})
                            </div>
                            <a href="/digital-indicator?job_no=${encodeURIComponent(q.toUpperCase())}" class="omni-item">
                                <div class="omni-item-left">
                                    <div class="omni-item-icon"><i class="fa-solid fa-microchip"></i></div>
                                    <div>
                                        <div class="omni-item-title">Open Digital Indicator with ${q.toUpperCase()}</div>
                                        <div class="omni-item-sub">Auto-fetch ST-ERP metadata &amp; multi-system data</div>
                                    </div>
                                </div>
                                <span class="omni-item-badge">Digital Indicator</span>
                            </a>
                            <a href="/belt-scale?job_no=${encodeURIComponent(q.toUpperCase())}" class="omni-item">
                                <div class="omni-item-left">
                                    <div class="omni-item-icon"><i class="fa-solid fa-truck-ramp-box"></i></div>
                                    <div>
                                        <div class="omni-item-title">Open Belt Scale with ${q.toUpperCase()}</div>
                                        <div class="omni-item-sub">Auto-fetch ST-ERP order details</div>
                                    </div>
                                </div>
                                <span class="omni-item-badge">Belt Scale</span>
                            </a>
                            <div style="height: 1px; background: #e2e8f0; margin: 6px 0;"></div>
                        `;
                    }

                    // Match systems
                    const matches = ALL_SYSTEMS.filter(s => 
                        s.name.toLowerCase().includes(q) || 
                        s.tags.toLowerCase().includes(q) || 
                        s.cat.toLowerCase().includes(q)
                    );

                    if (matches.length > 0) {
                        html += `
                            <div style="padding: 8px 12px 6px; font-size: 11px; font-weight: 800; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px;">
                                Equipment Systems (${matches.length})
                            </div>
                        `;
                        matches.slice(0, 6).forEach(m => {
                            html += `
                                <a href="${m.url}" class="omni-item">
                                    <div class="omni-item-left">
                                        <div class="omni-item-icon"><i class="fa-solid ${m.icon}"></i></div>
                                        <div>
                                            <div class="omni-item-title">${m.name}</div>
                                            <div class="omni-item-sub">${m.cat} System</div>
                                        </div>
                                    </div>
                                    <span class="omni-item-badge"><i class="fa-solid fa-arrow-right"></i> Open</span>
                                </a>
                            `;
                        });
                    }

                    if (!html) {
                        html = `
                            <div style="padding: 20px; text-align: center; color: #94a3b8; font-size: 13.5px;">
                                <i class="fa-solid fa-magnifying-glass" style="font-size: 20px; margin-bottom: 8px; display: block;"></i>
                                No matching system or job found for "<strong>${q}</strong>"
                            </div>
                        `;
                    }

                    dropdown.innerHTML = html;
                    dropdown.style.display = 'block';
                });

                // Close on click outside
                document.addEventListener('click', function(e) {
                    if (!searchInput.contains(e.target) && !dropdown.contains(e.target)) {
                        dropdown.style.display = 'none';
                    }
                });

                searchInput.addEventListener('keydown', function(e) {
                    if (e.key === 'Enter') {
                        e.preventDefault();
                        executeOmniSearch();
                    }
                });
            }

            window.executeOmniSearch = function() {
                const q = (document.getElementById('omniSearchInput').value || '').trim();
                if (!q) return;
                
                // If job number, default to digital indicator
                if (/^(sy|job|bs|di|cs|wf|ord|wo)/i.test(q) || /\\d{3,}/.test(q)) {
                    window.location.href = `/digital-indicator?job_no=${encodeURIComponent(q.toUpperCase())}`;
                    return;
                }

                // Match first system
                const qLower = q.toLowerCase();
                const m = ALL_SYSTEMS.find(s => s.name.toLowerCase().includes(qLower) || s.tags.includes(qLower));
                if (m) {
                    window.location.href = m.url;
                } else {
                    window.location.href = `/digital-indicator?job_no=${encodeURIComponent(q.toUpperCase())}`;
                }
            };
        });
    </script>
</body>
"""

body_end = final_html.rfind('</body>')
final_html_with_script = final_html[:body_end] + omni_script + final_html[body_end + 7:]

with open(portal_path, 'w', encoding='utf-8') as f:
    f.write(final_html_with_script)

print('SUCCESS: Updated portal.html with UI-UX Pro Max Bento Box Grid & Omni Search!')
