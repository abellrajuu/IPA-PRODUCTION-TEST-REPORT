import sys, os, time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

opts = Options()
opts.add_argument('--headless=new')
opts.add_argument('--window-size=1440,1024')
driver = webdriver.Chrome(options=opts)

out_dir = r"C:\Users\Admin\.gemini\antigravity\brain\290da693-b5ac-4320-a0b6-718e0ed3733d"

try:
    # 1. Login
    driver.get('http://localhost:5005/login')
    time.sleep(0.5)
    driver.find_element(By.NAME, 'username').send_keys('a')
    driver.find_element(By.NAME, 'password').send_keys('a')
    driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()
    time.sleep(1)

    # 2. Digital indicator form (top)
    driver.get('http://localhost:5005/digital-indicator')
    time.sleep(1)
    driver.save_screenshot(os.path.join(out_dir, 'ui_audit_digital_indicator.png'))
    print('Captured ui_audit_digital_indicator.png')

    # 3. Fill with SY0192 to show multi-system Sticky bar & cards
    job_input = driver.find_element(By.ID, 'job_no')
    job_input.clear()
    job_input.send_keys('SY0192')
    job_input.send_keys(Keys.ENTER)
    time.sleep(2)
    driver.save_screenshot(os.path.join(out_dir, 'ui_audit_sy0192_filled.png'))
    print('Captured ui_audit_sy0192_filled.png')

    # 4. Scroll down to show calibration matrix
    driver.execute_script('window.scrollTo(0, 1100);')
    time.sleep(0.5)
    driver.save_screenshot(os.path.join(out_dir, 'ui_audit_calibration_matrix.png'))
    print('Captured ui_audit_calibration_matrix.png')

    # 5. Review Dashboard
    driver.get('http://localhost:5005/review-dashboard')
    time.sleep(1)
    driver.save_screenshot(os.path.join(out_dir, 'ui_audit_review_dashboard.png'))
    print('Captured ui_audit_review_dashboard.png')

except Exception as e:
    print('Error:', e)
finally:
    driver.quit()
