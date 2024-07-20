import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import traceback
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
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


#鲁棒返回对象（绝对）非一定、可点
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
            
#鲁棒返回对象（绝对）非一定、可点
def wait_return_subelement_absolute_notmust_v1(driver, wait, time_w, xpath, times):
    i = 1
    while(1):
        try:
            target_element = WebDriverWait(driver, time_w).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            return target_element
        except:
            i+=1
            if i>=times:
                return None

            
#鲁棒返回对象（绝对）非一定、可视
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
            

def wait_return_element_two_xpath(driver, xpath1, xpath2, timewait):
    xpath1 = xpath1
    xpath2 = xpath2
    timeout = timewait
    element = None

    try:
        # 尝试使用第一个 XPath
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath1)))
        return element
    except (NoSuchElementException, TimeoutException):
        try:
            # 如果第一个 XPath 找不到，尝试使用第二个 XPath
            element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath2)))
            return element
        except (NoSuchElementException, TimeoutException):
            return False