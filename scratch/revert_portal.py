import sys, os

portal_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\portal.html'
with open(portal_path, 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Remove bento_css from <style>
bento_marker = '/* ══════════════════════════════════════════════════════════════════\n           UI-UX PRO MAX: BENTO BOX WORKSTATION & OMNI-SEARCH STYLES'
if bento_marker not in text:
    bento_marker = 'UI-UX PRO MAX: BENTO BOX WORKSTATION'

idx_css = text.find(bento_marker)
if idx_css != -1:
    css_start = text.rfind('\n', 0, idx_css)
    css_end = text.find('</style>', idx_css)
    text = text[:css_start] + '\n    ' + text[css_end:]
    print('SUCCESS: Removed bento_css from portal.html')
else:
    print('bento_marker not found in CSS')

# 2. Restore original view-home
og_view_home = """<div id="view-home" class="view active">
        <div class="workstation-header">
            <h1 class="workstation-title">Production Test Report</h1>
            <p class="workstation-sub">IPA Weighing &amp; Automation</p>
        </div>

        <div class="home-cards-container">
            
            <!-- CARD 1: TEST REPORTS -->
            <div class="hub-card hub-card-primary" onclick="openView('test-reports-hub')">
                <div class="card-icon">
                    <i class="fa-solid fa-flask-vial"></i>
                </div>
                <div class="card-title">Test Reports</div>
            </div>

            <!-- CARD 2: VIEW ALL REPORTS -->
            <div class="hub-card hub-card-archive" onclick="window.location.href='/review-dashboard'">
                <div class="card-icon">
                    <i class="fa-solid fa-database"></i>
                </div>
                <div class="card-title">View All Reports</div>
            </div>

        </div>
    </div>


    """

vh_start = text.find('<div id="view-home"')
vh_end = text.find('<div id="view-test-reports-hub"')
if vh_start != -1 and vh_end != -1:
    text = text[:vh_start] + og_view_home + text[vh_end:]
    print('SUCCESS: Restored original view-home in portal.html')
else:
    print('Error finding view-home or view-test-reports-hub')

# 3. Remove omni_script
omni_marker = '<!-- UI-UX PRO MAX: OMNI SEARCH & LIVE DATE SCRIPT -->'
idx_omni = text.find(omni_marker)
if idx_omni != -1:
    end_script = text.find('</body>', idx_omni)
    text = text[:idx_omni] + text[end_script:]
    print('SUCCESS: Removed omni_script from portal.html')
else:
    print('omni_marker not found')

with open(portal_path, 'w', encoding='utf-8') as f:
    f.write(text)

print('Portal.html restored to OG state!')
