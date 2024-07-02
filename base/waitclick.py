import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import traceback
from selenium.webdriver.support.ui import WebDriverWait
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


#鲁棒绝对点选非一定
def wait_click_xpath_notmust(wait, time_w, xpath):
    i = 1
    while(1):
        if i>=3:
            break
        try:
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            break
        except:
            time.sleep(time_w)
            i+=1
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
            
            button = WebDriverWait(element,time_w).until(EC.element_to_be_clickable((By.XPATH, xpath)))
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
            button = WebDriverWait(element,time_w).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            time.sleep(time_w)


#鲁棒返回对象（相对）
def wait_return_subelement_relative(time_w, element,xpath):
    while(1):
        try:
            target_element = WebDriverWait(element,time_w).until(EC.presence_of_element_located((By.XPATH, xpath)))
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


#返回对象文本（绝对）
def wait_return_element_text(wait, time_w, xpath):
    while(1):
        try:
            target_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            return target_element.get_attribute('textContent')
        except:
            traceback.print_exc()
            time.sleep(time_w)


#鲁棒返回对象（绝对）非一定
def wait_return_subelement_absolute_notmust(wait, time_w, xpath, times):
    i = 1
    while(1):
        try:
            target_element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            return target_element
        except:
            time.sleep(time_w)
            i+=1
            if i>=times:
                return None
            
#鲁棒返回对象（绝对）非一定
def wait_return_subelement_absolute_notmust_visiable(wait, time_w, xpath, times):
    i = 1
    while(1):
        try:
            target_element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            return target_element
        except:
            time.sleep(time_w)
            i+=1
            if i>=times:
                return None