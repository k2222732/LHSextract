from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
#import xlwings as xw
from openpyxl import load_workbook
from selenium import webdriver
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pandas as pd
import time 
import os
import requests
import re
import inspect



amount_that_complete = 0

directory = "g:/project/LHSextract/database/database_org"

def driver_create(chrome_path, chromedriver_path):
    chrome_options = Options()
    # 替换为你的 Chrome 浏览器的实际安装路径
    chrome_options.binary_location = chrome_path  
    # 创建 Service 对象并指定 ChromeDriver 的路径
    service = Service(executable_path=chromedriver_path)
    # 启动 WebDriver
    driver = webdriver.Chrome(service=service, options=chrome_options)
    #返回driver
    print(f"浏览器驱动创建成功")
    return driver


def login(account, password, driver, url, wait):
    while(1):
        # 打开网址
        driver.get(url)
        username_box = wait.until(EC.visibility_of_element_located((By.ID, 'username')))
        password_box = wait.until(EC.visibility_of_element_located((By.ID, 'password')))
        validatecode = wait.until(EC.visibility_of_element_located((By.ID, 'validateCode')))
        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.js-submit.tianze-loginbtn')))
        username_box.send_keys(Keys.CONTROL + "a")
        username_box.send_keys(Keys.BACKSPACE)
        username_box.send_keys(account)
        password_box.send_keys(Keys.CONTROL + "a")
        password_box.send_keys(Keys.BACKSPACE)
        password_box.send_keys(password)
        # 等待手动输入验证码
        temp = input ("Please enter the captcha and hit enter in the browser")
        validatecode.send_keys(Keys.CONTROL + "a")
        validatecode.send_keys(Keys.BACKSPACE)
        validatecode.send_keys(temp)
        # 点击登录按钮
        login_button.click()
        #等待1秒
        time.sleep(2)
        #获取当前网页的doom
        response = driver.page_source
        #检查doom里是否有"您上次登录是"字样
        if "您上次登录是" in response:
        #如果有打印登录成功，跳出循环
            break
        #如果没有继续本函数上面代码
        else:
            continue
    print(f"登录成功")


def access_org_database(driver, wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//img[contains(@src, "发展党员纪实公示系统.png")])[2]')
    try:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if '发展党员纪实公示系统' in driver.title:
                break
        print(f"进入发展党员纪实公示系统成功")
    except:
        time.sleep(1)


def switch_role(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[@class = "el-dropdown-link el-dropdown-selfdefine   "]')
    print(f"进入角色下拉列表成功")
    wait_click_xpath(wait, time_w = 0.5, xpath = '//li[contains(text(), "中国共产党山东汶上经济开发区工作委员会")]')
    print(f"切换角色成功")


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
        except:
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            time.sleep(time_w)


