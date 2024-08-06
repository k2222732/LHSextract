from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from datetime import datetime
import pandas as pd
import time 
import os
import re
import configparser
import threading
from bs4 import BeautifulSoup
from base.waitclick import *
from base.write_entry import *
from middle.dev_download import *
from openpyxl import load_workbook
import openpyxl
from base.validate import *

from middle.dev_middle import *
stop_event = threading.Event()
#dev_func.py进入个人页面的操作
def enter_person_infopage(wait, driver0, row_number, member_excel, member_excel_path, page_number, control, temp_countx, xpath, wb, ws):
    #//li[@class = 'el-select-dropdown__item selected hover']//span
    value = ""
    if control == 1:
        value = "正式党员"
    elif control == 2:
        value = "预备党员"
    elif control == 3:
        value = "发展对象"
    elif control == 4:
        value = "入党积极分子"
    elif control == 5:
        value = "入党申请人"
    while 1:
        result =validate_list_content(wait, xpath, value=value)
        if result == True:
            access_info_page(wait, row_number, file = member_excel, path = member_excel_path, temp_countx=temp_countx,wb=wb, ws=ws)
            try:
                for handle in driver0.window_handles:
                    driver0.switch_to.window(handle)
                    if '入党申请人基本信息' in driver0.page_source:
                        break
            except:
                time.sleep(0.3)
            print(f"进入入党申请人基本信息页面成功")
            break
        elif result == False:
            if value == "正式党员":
                switch_formal_mem(wait)
            elif value == "预备党员":
                switch_informal_mem(wait)
            elif value == "发展对象":
                switch_devtarg(wait)
            elif value == "入党积极分子":
                switch_activist(wait)
            elif value == "入党申请人":
                switch_applicant(wait)
            set_amount_perpage(wait)
            input_page = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            input_page.send_keys(Keys.CONTROL + "a")
            input_page.send_keys(Keys.BACKSPACE)
            input_page.send_keys(page_number)
            input_page.send_keys(Keys.RETURN)
            time.sleep(0.5)
            

