import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

def wait_click_xpath(wait, time_w, xpath):
    while(1):
        try:
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            break
        except:
            time.sleep(time_w)
    while(1):
        try:
            button.click()
            break
        except Exception as e:
            print(f"An exception occurred: {e}")
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            time.sleep(time_w)