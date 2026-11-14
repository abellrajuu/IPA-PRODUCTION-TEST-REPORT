import sys, os

dash_path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\frontend_html_templates\review_dashboard.html'
with open(dash_path, 'r', encoding='utf-8') as f:
    text = f.read()

# Let's inspect the filter tabs markup
idx = text.find('<div class="filter-tabs">')
print(text[idx-20:idx+600])
