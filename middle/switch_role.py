import time
from base.waitclick import  wait_click_xpath
from selenium.webdriver.support import expected_conditions as EC
import configparser
from selenium.webdriver.common.by import By


def access_e_shandong(driver, wait):
    global driver0
    driver0 = driver
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//img[contains(@src, "山东e支部.png")])[2]')
    time.sleep(0.8)
    try:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if '山东e支部管理系统' in driver.title:
                break
    except:
        time.sleep(1)



def switch_role(wait):
    while 1:
        try:
            droplist = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.avatar-wrapper.fs-dropdown-selfdefine')))
            droplist.click()
            print(f"进入角色下拉列表成功")
            time.sleep(1)
            config_file = 'config.ini'
            config = configparser.ConfigParser()
            config.read(config_file)
            role_name_mem = config.get('role_e_name', 'name_role_e', fallback='')
            role = wait.until(EC.element_to_be_clickable((By.XPATH, f"//SPAN[contains(text(), '{role_name_mem}')]")))
            role.click()
            if wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.avatar-wrapper.fs-dropdown-selfdefine'))).text == f"{role_name_mem}":
                break
            else:
                pass
        except:
            time.sleep(1)