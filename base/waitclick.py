import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import traceback

#鲁棒绝对点选
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
        except Exception:
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            time.sleep(time_w)


#鲁棒相对点选
def wait_click_xpath_relative(wait, time_w, element,xpath):
    while(1):
        try:
            button = element.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            break
        except:
            traceback.print_exc()
            time.sleep(time_w)
    while(1):
        try:
            button.click()
            break
        except Exception:
            traceback.print_exc()
            button = element.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            time.sleep(time_w)


#鲁棒返回对象（相对）
def wait_return_subelement_relative(wait, time_w, element,xpath):
    while(1):
        try:
            target_element = element.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            return target_element
        except:
            traceback.print_exc()
            time.sleep(time_w)


#鲁棒返回对象（绝对）
def wait_return_subelement_absolute(wait, time_w, xpath):
    while(1):
        try:
            target_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            return target_element
        except:
            traceback.print_exc()
            time.sleep(time_w)