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
    # 1. Login
    driver.get('http://localhost:5005/login')
    time.sleep(0.5)
    driver.find_element(By.NAME, 'username').send_keys('a')
    driver.find_element(By.NAME, 'password').send_keys('a')
    driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()
    time.sleep(1)

    # 2. Portal: type SY0192 into Omni-Search
    search_input = driver.find_element(By.ID, 'omniSearchInput')
    search_input.send_keys('SY0192')
    time.sleep(0.6)

    # Click the first suggestion item
    first_suggestion = driver.find_element(By.CSS_SELECTOR, '.omni-item')
    print('Clicking suggestion:', first_suggestion.text.replace('\n', ' -- '))
    first_suggestion.click()

    # 3. Wait for the digital indicator form to auto-load with SY0192
    time.sleep(2.5)
    driver.save_screenshot(os.path.join(out_dir, 'ui_pro_max_sy0192_autolaunch.png'))
    print('Captured ui_pro_max_sy0192_autolaunch.png')

    # Verify fields populated
    cust_val = driver.find_element(By.ID, 'customer').get_attribute('value')
    cap_val = driver.find_element(By.ID, 'capacity').get_attribute('value')
    print(f'Customer value: "{cust_val}", Capacity: "{cap_val}"')

    # Check Step 4
    driver.execute_script('window.scrollTo(0, 1000);')
    time.sleep(0.5)
    driver.save_screenshot(os.path.join(out_dir, 'ui_pro_max_sy0192_autolaunch_step4.png'))
    print('Captured ui_pro_max_sy0192_autolaunch_step4.png')

except Exception as e:
    print('Error in flow:', e)
finally:
    driver.quit()
