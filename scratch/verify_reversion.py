import sys, os, time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

opts = Options()
opts.add_argument('--headless=new')
opts.add_argument('--window-size=1440,1024')
driver = webdriver.Chrome(options=opts)

out_dir = r"C:\Users\Admin\.gemini\antigravity\brain\290da693-b5ac-4320-a0b6-718e0ed3733d"

try:
    driver.get('http://localhost:5005/login')
    time.sleep(0.5)
    driver.find_element(By.NAME, 'username').send_keys('a')
    driver.find_element(By.NAME, 'password').send_keys('a')
    driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()
    time.sleep(1)

    driver.save_screenshot(os.path.join(out_dir, 'ui_restored_portal.png'))
    print('Captured ui_restored_portal.png')

    driver.get('http://localhost:5005/review-dashboard')
    time.sleep(1)
    driver.save_screenshot(os.path.join(out_dir, 'ui_restored_review_dashboard.png'))
    print('Captured ui_restored_review_dashboard.png')

    driver.get('http://localhost:5005/digital-indicator')
    time.sleep(1)
    driver.save_screenshot(os.path.join(out_dir, 'ui_restored_digital_indicator.png'))
    print('Captured ui_restored_digital_indicator.png')

except Exception as e:
    print('Error:', e)
finally:
    driver.quit()
