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
from base.waitclick import wait_return_subelement_absolute
stop_event = threading.Event()

def stop_dev_thread():
    stop_event.set()




captcha = ""
amount_mem_complete = 0
amount_infomem_complete = 0
amount_activist_complete = 0
amount_devtar_complete = 0
amount_applicant_complete = 0

mem_total_amount = 0
infomem_total_amount = 0
activist_total_amount = 0
devtar_total_amount = 0
applicant_total_amount = 0

dev_directory = ""

driver0 = ""
temp_countx = 0

def access_dev_database(driver, wait):
    global driver0
    driver0 = driver
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//img[contains(@src, "发展党员纪实公示系统.png")])[2]')
    try:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if '发展党员纪实公示系统' in driver.title:
                break
        print(f"进入发展党员纪实公示系统成功")
    except:
        time.sleep(1)


def switch_role(wait, driver):
    while 1:
        try:
            wait_click_xpath(wait, time_w = 0.5, xpath = "//i[@class = 'el-icon-caret-bottom']")
            print(f"进入角色下拉列表成功")
            config_file = 'config.ini'
            config = configparser.ConfigParser()
            config.read(config_file)
            role_name_dev = config.get('role_dev_name', 'name_role_dev', fallback='')
            wait_click_xpath_action(driver, wait, time_w = 0.5, xpath = f'//li[contains(text(), "{role_name_dev}")]')
            print(f"切换角色成功")
            break
        except:
            time.sleep(2)


#切换到正式党员信息页面
def switch_formal_mem(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]//li//span[contains(text(), "正式党员")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')
    
def switch_informal_mem(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]//li//span[contains(text(), "预备党员")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')

def switch_devtarg(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]//li//span[contains(text(), "发展对象")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')

def switch_activist(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]//li//span[contains(text(), "入党积极分子")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')

def switch_applicant(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[2]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]//li//span[contains(text(), "入党申请人")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')






def new_excel(wait):
    global mem_total_amount
    global infomem_total_amount
    global devtar_total_amount
    global activist_total_amount
    global applicant_total_amount
    global dev_directory
    config = configparser.ConfigParser()
    config.read('config.ini')
    dev_directory = config.get('Paths_dev_info', 'dev_info_path')
    today_date = datetime.now().strftime("%Y-%m-%d")
    excel_file_name = f"{today_date}党员发展纪实信息库.xlsx"
    excel_file_path = os.path.join(dev_directory, excel_file_name)
    #if os.path.isfile(excel_file_path):
        #member_excel = load_workbook(excel_file_path)
        #rebuild(excel_file_path, wait, member_total_amount, member_excel, excel_file_path)
    #else:
        #在当前文件夹新建一个文件夹命名为"database"在里面新建一个文件夹名为"database_member"
        #os.makedirs(dev_directory, exist_ok=True)
        #在database_member文件夹里新建一个excel文件命名为"今天的日期"&"党员信息库"

        #在excel里的第一行建立表头，分别是"姓名、性别、公民身份号码、民族、出生日期、学历、人员类别、学位、所在党支部、手机号码、入党日期
        #转正日期、党龄、党龄校正值、新社会阶层类型、工作岗位、从事专业技术职务、是否农民工、现居住地、户籍所在地、是否失联党员、是否流动党员、
        #党员注册时间、注册手机号、党员增加信息、党员减少信息、入党类型、转正情况、入党时所在支部、延长预备期时间
    if os.path.isfile(excel_file_path):
        member_excel = load_workbook(excel_file_path)
        rebuild(excel_file_path, wait, member_excel)
    else:
        os.makedirs(dev_directory, exist_ok=True)
        columns = ["序号", "姓名","性别","公民身份证号码","民族","出生日期","学历","申请入党日期","手机号码","背景信息","工作岗位",
                "政治面貌","接受申请党组织","籍贯","入团日期","参加工作日期","申请入党日期","确定入党积极分子日期",
                "工作单位及职务","家庭住址","加入党组织日期","转为正式党员日期","人员类别","党籍状态","入党类型","所在党支部", "人员类型", "隶属党组织"]
            
        df = pd.DataFrame(columns=columns)
        df.to_excel(excel_file_path, index=False)
        member_excel = load_workbook(excel_file_path)
        print(f"文件 '{excel_file_path}' 已成功创建。")
        switch_formal_mem(wait)
        mem_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
        synchronizing(wait, member_excel, excel_file_path, control=1)
        if stop_event.is_set():
            return
        switch_informal_mem(wait)
        infomem_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
        synchronizing(wait, member_excel, excel_file_path, control=2)
        if stop_event.is_set():
            return
        switch_devtarg(wait)
        temp = wait.until(EC.presence_of_element_located((By.ID, "tabls")))
        temp_html = temp.get_attribute("outerHTML")
        if "暂无数据" in temp_html:
            pass
        else:
            devtar_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
            synchronizing(wait, member_excel, excel_file_path, control=3)
        if stop_event.is_set():
            return
        switch_activist(wait)
        activist_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
        synchronizing(wait, member_excel, excel_file_path, control=4)
        if stop_event.is_set():
            return
        switch_applicant(wait)
        applicant_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
        synchronizing(wait, member_excel, excel_file_path, control=5)
        if stop_event.is_set():
            return


def rebuild(excel_file_path, wait, member_excel):
    global mem_total_amount
    global infomem_total_amount
    global devtar_total_amount
    global activist_total_amount
    global applicant_total_amount

    global amount_mem_complete
    global amount_infomem_complete
    global amount_activist_complete
    global amount_devtar_complete
    global amount_applicant_complete

    switch_formal_mem(wait)
    mem_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")  
    switch_informal_mem(wait)
    infomem_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
    switch_devtarg(wait)
    temp = wait.until(EC.presence_of_element_located((By.ID, "tabls")))
    temp_html = temp.get_attribute("outerHTML")
    if "暂无数据" in temp_html:
        devtar_total_amount = 0
    else:
        devtar_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
    switch_activist(wait)
    activist_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
    switch_applicant(wait)
    applicant_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
    #先改变全局变量
    a = init_complete_amount(excel_file_path)
    if a <= mem_total_amount:
        amount_mem_complete = a
        synchronizing(wait, member_excel, excel_file_path, control = 1)
        if stop_event.is_set():
            return
        switch_informal_mem(wait)
        synchronizing(wait, member_excel, excel_file_path, control=2)
        if stop_event.is_set():
            return
        switch_devtarg(wait)
        temp = wait.until(EC.presence_of_element_located((By.ID, "tabls")))
        temp_html = temp.get_attribute("outerHTML")
        if "暂无数据" in temp_html:
            pass
        else:
            devtar_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
        synchronizing(wait, member_excel, excel_file_path, control=3)
        if stop_event.is_set():
            return
        switch_activist(wait)
        synchronizing(wait, member_excel, excel_file_path, control=4)
        if stop_event.is_set():
            return
        switch_applicant(wait)
        synchronizing(wait, member_excel, excel_file_path, control=5)
        if stop_event.is_set():
            return
    elif a>mem_total_amount and a <= mem_total_amount + infomem_total_amount:
        amount_mem_complete = mem_total_amount
        amount_infomem_complete = a - mem_total_amount
        synchronizing(wait, member_excel, excel_file_path, control = 2)
        if stop_event.is_set():
            return
        switch_devtarg(wait)
        temp = wait.until(EC.presence_of_element_located((By.ID, "tabls")))
        temp_html = temp.get_attribute("outerHTML")
        if "暂无数据" in temp_html:
            pass
        else:
            devtar_total_amount = get_total_amount(wait, xpath = "//span[@class = 'el-pagination__total']")
        synchronizing(wait, member_excel, excel_file_path, control=3)
        if stop_event.is_set():
            return
        switch_activist(wait)
        synchronizing(wait, member_excel, excel_file_path, control=4)
        if stop_event.is_set():
            return
        switch_applicant(wait)
        synchronizing(wait, member_excel, excel_file_path, control=5)
        if stop_event.is_set():
            return
    elif a>mem_total_amount + infomem_total_amount and a <= mem_total_amount + infomem_total_amount + devtar_total_amount:
        amount_mem_complete = mem_total_amount
        amount_infomem_complete = infomem_total_amount
        amount_devtar_complete = a - (mem_total_amount + infomem_total_amount)
        synchronizing(wait, member_excel, excel_file_path, control = 3)
        if stop_event.is_set():
            return
        switch_activist(wait)
        synchronizing(wait, member_excel, excel_file_path, control=4)
        if stop_event.is_set():
            return
        switch_applicant(wait)
        synchronizing(wait, member_excel, excel_file_path, control=5)
        if stop_event.is_set():
            return
    elif a>mem_total_amount + infomem_total_amount + devtar_total_amount and a <= mem_total_amount + infomem_total_amount + devtar_total_amount + activist_total_amount:
        amount_mem_complete = mem_total_amount
        amount_infomem_complete = infomem_total_amount
        amount_devtar_complete = devtar_total_amount
        amount_activist_complete = a - (mem_total_amount + infomem_total_amount + devtar_total_amount)
        synchronizing(wait, member_excel, excel_file_path, control = 4)
        if stop_event.is_set():
            return
        switch_applicant(wait)
        synchronizing(wait, member_excel, excel_file_path, control=5)
        if stop_event.is_set():
            return
    elif a>mem_total_amount + infomem_total_amount + devtar_total_amount + activist_total_amount and a <= mem_total_amount + infomem_total_amount + devtar_total_amount + activist_total_amount + applicant_total_amount :
        amount_mem_complete = mem_total_amount
        amount_infomem_complete = infomem_total_amount
        amount_devtar_complete = devtar_total_amount
        amount_activist_complete = activist_total_amount
        amount_applicant_complete = a - (mem_total_amount + infomem_total_amount + devtar_total_amount + activist_total_amount)
        synchronizing(wait, member_excel, excel_file_path, control = 5)
        if stop_event.is_set():
            return

def init_complete_amount(excel_file_path):
    df = pd.read_excel(excel_file_path, sheet_name=0)
    row_count = df.dropna(how='all').shape[0]
    amount_that_complete = row_count - 1
    return amount_that_complete




def synchronizing(wait, member_excel, member_excel_path, control):

    global amount_mem_complete
    global amount_infomem_complete
    global amount_devtar_complete
    global amount_activist_complete
    global amount_applicant_complete

    global mem_total_amount
    global infomem_total_amount
    global devtar_total_amount
    global activist_total_amount
    global applicant_total_amount

    
    
    if control == 1:
        schedule(complete = amount_mem_complete, total = mem_total_amount, xpath = "//input[@type = 'number']", wait = wait, member_excel = member_excel, member_excel_path = member_excel_path, control = control)
    elif control == 2:
        schedule(complete = amount_infomem_complete, total = infomem_total_amount, xpath = "//input[@type = 'number']", wait = wait, member_excel = member_excel, member_excel_path = member_excel_path, control = control)
    elif control == 3:
        schedule(complete = amount_devtar_complete, total = devtar_total_amount, xpath = "//input[@type = 'number']", wait = wait, member_excel = member_excel, member_excel_path = member_excel_path, control = control)
    elif control == 4:
        schedule(complete = amount_activist_complete, total = activist_total_amount, xpath = "//input[@type = 'number']", wait = wait, member_excel = member_excel, member_excel_path = member_excel_path, control = control)
    elif control == 5:
        schedule(complete = amount_applicant_complete, total = applicant_total_amount, xpath = "//input[@type = 'number']", wait = wait, member_excel = member_excel, member_excel_path = member_excel_path, control = control)


def get_total_amount(wait, xpath):
    time.sleep(1)
    try:
        char_amountof_member = wait.until(EC.presence_of_element_located((By.XPATH, xpath))).text
        int_amountof_members = re.findall(r'\d+', char_amountof_member)
        int_amountof_member = int(int_amountof_members[0]) if int_amountof_members else None
        return int(int_amountof_member)
    except:
        #//div[@class = "el-table y_table el-table--fit el-table--border el-table--enable-row-hover el-table--enable-row-transition"]
        can = wait_return_subelement_absolute(wait, 0.5, "//main[@class ='el-main main']")
        can_html = can.get_attribute('outerHTML')
        soup = BeautifulSoup(can_html, 'html.parser')
        result = soup.find('span', class_ = "el-table__empty-text")
        if result:
            return 0
        else:
            return -1



def set_amount_perpage(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = "(//input[@class = 'el-input__inner'])[4]")
    wait_click_xpath(wait, time_w = 0.5, xpath = "//span[contains(text(), '100条/页')]")
    



def schedule(complete, total, xpath, wait, member_excel, member_excel_path, control):
    global driver0
    global amount_mem_complete
    global amount_infomem_complete
    global amount_activist_complete
    global amount_devtar_complete
    global amount_applicant_complete

    if control == 1:
        switch_formal_mem(wait)
    elif control == 2:
        switch_informal_mem(wait)
    elif control == 3:
        switch_devtarg(wait)
    elif control == 4:
        switch_activist(wait)
    elif control == 5:
        switch_applicant(wait)

    set_amount_perpage(wait)
    
    while complete < total:
        page_number = int(complete / 100 + 1)
        row_number = int(complete % 100 + 1)
        #time.sleep(0.1)
        input_page = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        input_page.send_keys(Keys.CONTROL + "a")
        input_page.send_keys(Keys.BACKSPACE)
        input_page.send_keys(page_number)
        input_page.send_keys(Keys.RETURN)
        access_info_page(wait, row_number, file = member_excel, path = member_excel_path)
        try:
            for handle in driver0.window_handles:
                driver0.switch_to.window(handle)
                if '入党申请人基本信息' in driver0.page_source:
                    break
        except:
            time.sleep(0.3)
        print(f"进入入党申请人基本信息页面成功")
        downloading(file = member_excel, wait = wait, path = member_excel_path, control = control)
        driver0.close()
        try:
            for handle in driver0.window_handles:
                driver0.switch_to.window(handle)
                if '发展党员纪实' in driver0.title:
                    break
        except:
            time.sleep(0.3)

        ##在这里检查线程关闭信号
        if stop_event.is_set():
            break

        if control == 1:
            complete = amount_mem_complete
        elif control == 2:
            complete = amount_infomem_complete
        elif control == 3:
            complete = amount_devtar_complete
        elif control == 4:
            complete = amount_activist_complete
        elif control == 5:
            complete = amount_applicant_complete



def access_info_page(wait, rowx, file, path):
    global temp_countx
    lishudangzuzhi = wait.until(EC.visibility_of_element_located((By.XPATH, f"//table[@class = 'el-table__body']/tbody/tr[{rowx}]/td[5]//a"))).text
    file.active.cell(row=temp_countx+2, column=28).value = lishudangzuzhi
    file.save(path)
    wait_click_xpath(wait, time_w = 0.5, xpath = f"//table[@class = 'el-table__body']/tbody/tr[{rowx}]/td[1]//a")
    
    
    


def bitian_located(file, path, row, column, wait, str):
    file.active.cell(row=row, column=column).value = wait.until(EC.presence_of_element_located((By.XPATH, str))).text
    start_time = time.time()
    while(1):
            elapsed_time = time.time() - start_time
            if elapsed_time > 5:
                break
            try:
                df = file.active.cell(row=row, column=column).value
                assert df != ""
                break
            except AssertionError:
                file.active.cell(row=row, column=column).value = wait.until(EC.presence_of_element_located((By.XPATH, str))).text
                file.save(path)


def downloading(file, wait, path, control):
    global driver0
    global amount_mem_complete
    global amount_infomem_complete
    global amount_activist_complete
    global amount_devtar_complete
    global amount_applicant_complete
    global temp_countx
    if control == 1:
        count = amount_mem_complete
    elif control == 2:
        count = amount_infomem_complete + amount_mem_complete
    elif control == 3:
        count = amount_activist_complete + amount_infomem_complete + amount_mem_complete
    elif control == 4:
        count = amount_devtar_complete + amount_activist_complete + amount_infomem_complete + amount_mem_complete
    elif control == 5:
        count = amount_applicant_complete + amount_devtar_complete + amount_activist_complete + amount_infomem_complete + amount_mem_complete

    if control == 1:
        count = count + 1
        countx = count + 1
        temp_countx = count
        #填写序号
        file.active.cell(row=countx, column=1).value = count
        file.save(path)
        
        #姓名
        name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text



        file.active.cell(row=countx, column=2).value = name_temp
        file.save(path)
        #循环断言
        while(1):
            try:
                df = file.active.cell(row=countx, column=2).value
                assert df != ""
                break
            except AssertionError:
                time.sleep(0.5)
                file.active.cell(row=countx, column=2).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
                file.save(path)
        #性别
        file.active.cell(row=countx, column=3).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[2]//span[1]"))).text
        
        #公民身份证号码
        file.active.cell(row=countx, column=4).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[4]//span[1]"))).text
        
        #民族
        file.active.cell(row=countx, column=5).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[6]//span[1]"))).text
        
        #出生日期
        file.active.cell(row=countx, column=6).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[4]//span[1]"))).text
        
        #学历
        bitian_located(file, path, countx, column=7, wait = wait, str = "(//tbody)[1]/tr[2]/td[6]//span[1]")
        


        #申请入党日期
        file.active.cell(row=countx, column=8).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[2]//span[1]"))).text
        
        #手机号码
        file.active.cell(row=countx, column=9).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[4]//span[1]"))).text
        
        #背景信息
        bitian_located(file, path, countx, column=10, wait = wait, str = "(//tbody)[1]/tr[3]/td[6]//span[1]")
        


        #工作岗位
        bitian_located(file, path, countx, column=11, wait = wait, str = "(//tbody)[1]/tr[4]/td[2]//span[1]")
        


        #政治面貌
        bitian_located(file, path, countx, column=12, wait = wait, str = "(//tbody)[1]/tr[4]/td[4]//span[1]")
        


        #接受申请党组织
        file.active.cell(row=countx, column=13).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[5]/td[2]//span[1]"))).text
        file.save(path)
        #切换到积极分子信息
        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
        #籍贯
        while 1:
            jiguan = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[2]/tr[3]/td[4]//span[1]"))).text
            if jiguan == "":
                pass
            else:
                file.active.cell(row=countx, column=14).value = jiguan
                break
        #入团日期
        file.active.cell(row=countx, column=15).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[4]//span)[1]"))).text
        
        #参加工作日期
        file.active.cell(row=countx, column=16).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[6]//span)[1]"))).text
        
        #申请入党日期
        file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[2]//span)[1]"))).text
        
        #确定入党积极分子日期
        file.active.cell(row=countx, column=18).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[4]//span)[1]"))).text
        
        #工作单位及职务
        file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[6]/td[4]//span)[1]"))).text
        
        #家庭住址
        file.active.cell(row=countx, column=20).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[7]/td[2]//span)[1]"))).text
        file.save(path)
        #切换到预备党员
        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '预备党员')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '预备党员')]")))
        # 加入党组织日期
        file.active.cell(row=countx, column=21).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[4]/td[6]"))).text
        
        # 转为正式党员日期
        file.active.cell(row=countx, column=22).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[5]/td[2]"))).text
               
        # 人员类别
        file.active.cell(row=countx, column=23).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[6]/td[2]"))).text
               
        # 党籍状态
        file.active.cell(row=countx, column=24).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[6]/td[4]"))).text
           
        # 入党类型
        file.active.cell(row=countx, column=25).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[6]/td[6]"))).text
        
        # 所在党支部
        file.active.cell(row=countx, column=26).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[8]/td[4]"))).text
        
        # 人员类型
        file.active.cell(row=countx, column=27).value = "正式党员"
        file.save(path)
        print("填写第",count,"个正式党员",name_temp,"信息成功") 
        amount_mem_complete = amount_mem_complete + 1



    elif control == 2:
        count = count + 1
        countx = count + 1
        temp_countx = count
        #填写序号
        file.active.cell(row=countx, column=1).value = count
        
        #姓名
        name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
        file.active.cell(row=countx, column=2).value = name_temp
        file.save(path)
        #循环断言
        while(1):
            try:
                df = file.active.cell(row=countx, column=2).value
                assert df != ""
                break
            except AssertionError:
                time.sleep(0.5)
                file.active.cell(row=countx, column=2).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
                file.save(path)
        #性别
        file.active.cell(row=countx, column=3).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[2]//span[1]"))).text
        
        #公民身份证号码
        file.active.cell(row=countx, column=4).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[4]//span[1]"))).text
        
        #民族
        file.active.cell(row=countx, column=5).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[6]//span[1]"))).text
        
        #出生日期
        file.active.cell(row=countx, column=6).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[4]//span[1]"))).text
       
        #学历
        bitian_located(file, path, countx, column=7, wait = wait, str = "(//tbody)[1]/tr[2]/td[6]//span[1]")


        #申请入党日期
        file.active.cell(row=countx, column=8).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[2]//span[1]"))).text
        
        #手机号码
        file.active.cell(row=countx, column=9).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[4]//span[1]"))).text
        
        #背景信息
        bitian_located(file, path, countx, column=10, wait = wait, str = "(//tbody)[1]/tr[3]/td[6]//span[1]")

        #工作岗位
        bitian_located(file, path, countx, column=11, wait = wait, str = "(//tbody)[1]/tr[4]/td[2]//span[1]")

        #政治面貌
        bitian_located(file, path, countx, column=12, wait = wait, str = "(//tbody)[1]/tr[4]/td[4]//span[1]")

        #接受申请党组织
        file.active.cell(row=countx, column=13).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[5]/td[2]//span[1]"))).text
        file.save(path)
        #切换到积极分子信息
        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
        #籍贯
        file.active.cell(row=countx, column=14).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[2]/tr[3]/td[4]//span[1]"))).text
      
        #入团日期
        file.active.cell(row=countx, column=15).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[4]//span)[1]"))).text
        
        #参加工作日期
        file.active.cell(row=countx, column=16).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[6]//span)[1]"))).text
        
        #申请入党日期
        file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[2]//span)[1]"))).text
        
        #确定入党积极分子日期
        file.active.cell(row=countx, column=18).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[4]//span)[1]"))).text
        
        #工作单位及职务
        file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[6]/td[4]//span)[1]"))).text
        
        #家庭住址
        file.active.cell(row=countx, column=20).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[7]/td[2]//span)[1]"))).text
        file.save(path)
        #切换到预备党员
        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '预备党员')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '预备党员')]")))
        # 加入党组织日期
        file.active.cell(row=countx, column=21).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[4]/td[6]"))).text
        
        # 转为正式党员日期
        file.active.cell(row=countx, column=22).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[5]/td[2]"))).text
              
        # 人员类别
        file.active.cell(row=countx, column=23).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[6]/td[2]"))).text
               
        # 党籍状态
        file.active.cell(row=countx, column=24).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[6]/td[4]"))).text
              
        # 入党类型
        file.active.cell(row=countx, column=25).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[6]/td[6]"))).text
        
        # 所在党支部
        file.active.cell(row=countx, column=26).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[4]/tr[8]/td[4]"))).text
        
        # 人员类型
        file.active.cell(row=countx, column=27).value = "预备党员"
        file.save(path)
        print("填写第",count,"个预备党员",name_temp,"信息成功") 
        amount_infomem_complete = amount_infomem_complete + 1



    elif control == 3:
        count = count + 1
        countx = count + 1
        temp_countx = count
        #填写序号
        file.active.cell(row=countx, column=1).value = count
        
        #姓名
        name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
        file.active.cell(row=countx, column=2).value = name_temp
        file.save(path)
        #循环断言
        while(1):
            try:
                df = file.active.cell(row=countx, column=2).value
                assert df != ""
                break
            except AssertionError:
                time.sleep(0.5)
                file.active.cell(row=countx, column=2).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
                file.save(path)
        #性别
        file.active.cell(row=countx, column=3).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[2]//span[1]"))).text
        
        #公民身份证号码
        file.active.cell(row=countx, column=4).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[4]//span[1]"))).text
        
        #民族
        file.active.cell(row=countx, column=5).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[6]//span[1]"))).text
        
        #出生日期
        file.active.cell(row=countx, column=6).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[4]//span[1]"))).text
        
        #学历
        bitian_located(file, path, countx, column=7, wait = wait, str = "(//tbody)[1]/tr[2]/td[6]//span[1]")

        #申请入党日期
        file.active.cell(row=countx, column=8).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[2]//span[1]"))).text
        
        #手机号码
        file.active.cell(row=countx, column=9).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[4]//span[1]"))).text
        
        #背景信息
        bitian_located(file, path, countx, column=10, wait = wait, str = "(//tbody)[1]/tr[3]/td[6]//span[1]")
        
        #工作岗位
        bitian_located(file, path, countx, column=11, wait = wait, str = "(//tbody)[1]/tr[4]/td[2]//span[1]")

        #政治面貌
        bitian_located(file, path, countx, column=12, wait = wait, str = "(//tbody)[1]/tr[4]/td[4]//span[1]")


        #接受申请党组织
        file.active.cell(row=countx, column=13).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[5]/td[2]//span[1]"))).text
        file.save(path)
        #切换到积极分子信息
        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
        #籍贯
        file.active.cell(row=countx, column=14).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[2]/tr[3]/td[4]//span[1]"))).text
        
        #入团日期
        file.active.cell(row=countx, column=15).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[4]//span)[1]"))).text
        
        #参加工作日期
        file.active.cell(row=countx, column=16).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[6]//span)[1]"))).text
        
        #申请入党日期
        file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[2]//span)[1]"))).text
        
        #确定入党积极分子日期
        file.active.cell(row=countx, column=18).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[4]//span)[1]"))).text
        
        #工作单位及职务
        file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[6]/td[4]//span)[1]"))).text
        
        #家庭住址
        file.active.cell(row=countx, column=20).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[7]/td[2]//span)[1]"))).text
        
        # 加入党组织日期
        file.active.cell(row=countx, column=21).value = "-"
        
        # 转为正式党员日期
        file.active.cell(row=countx, column=22).value = "-"
             
        # 人员类别
        file.active.cell(row=countx, column=23).value = "-"
             
        # 党籍状态
        file.active.cell(row=countx, column=24).value = "-"
             
        # 入党类型
        file.active.cell(row=countx, column=25).value = "-"
        
        # 所在党支部
        file.active.cell(row=countx, column=26).value = "-"
        file.save(path)
        # 人员类型
        file.active.cell(row=countx, column=27).value = "发展对象"
        print("填写第",count,"个发展对象",name_temp,"信息成功") 
        amount_devtar_complete = amount_devtar_complete + 1



    elif control == 4:
        count = count + 1
        countx = count + 1
        temp_countx = count
        #填写序号
        file.active.cell(row=countx, column=1).value = count
        
        #姓名
        name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
        file.active.cell(row=countx, column=2).value = name_temp
        file.save(path)
        #循环断言
        while(1):
            try:
                df = file.active.cell(row=countx, column=2).value
                assert df != ""
                break
            except AssertionError:
                time.sleep(0.5)
                file.active.cell(row=countx, column=2).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
                file.save(path)
        #性别
        file.active.cell(row=countx, column=3).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[2]//span[1]"))).text
        
        #公民身份证号码
        file.active.cell(row=countx, column=4).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[4]//span[1]"))).text
        
        #民族
        file.active.cell(row=countx, column=5).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[6]//span[1]"))).text
        
        #出生日期
        file.active.cell(row=countx, column=6).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[4]//span[1]"))).text
        
        #学历
        bitian_located(file, path, countx, column=7, wait = wait, str = "(//tbody)[1]/tr[2]/td[6]//span[1]")

        #申请入党日期
        file.active.cell(row=countx, column=8).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[2]//span[1]"))).text
        
        #手机号码
        file.active.cell(row=countx, column=9).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[4]//span[1]"))).text
        
        #背景信息
        bitian_located(file, path, countx, column=10, wait = wait, str = "(//tbody)[1]/tr[3]/td[6]//span[1]")

        #工作岗位
        bitian_located(file, path, countx, column=11, wait = wait, str = "(//tbody)[1]/tr[4]/td[2]//span[1]")

        #政治面貌
        bitian_located(file, path, countx, column=12, wait = wait, str = "(//tbody)[1]/tr[4]/td[4]//span[1]")


        #接受申请党组织
        file.active.cell(row=countx, column=13).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[5]/td[2]//span[1]"))).text
        file.save(path)
        #切换到积极分子信息

        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
        #籍贯
        file.active.cell(row=countx, column=14).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[2]/tr[3]/td[4]//span[1]"))).text
        
        #入团日期
        file.active.cell(row=countx, column=15).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[4]//span)[1]"))).text
        
        #参加工作日期
        file.active.cell(row=countx, column=16).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[6]//span)[1]"))).text
        
        #申请入党日期
        file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[2]//span)[1]"))).text
        
        #确定入党积极分子日期
        file.active.cell(row=countx, column=18).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[4]//span)[1]"))).text
       
        #工作单位及职务
        file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[6]/td[4]//span)[1]"))).text
       
        #家庭住址
        file.active.cell(row=countx, column=20).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[7]/td[2]//span)[1]"))).text
        
        # 加入党组织日期
        file.active.cell(row=countx, column=21).value = "-"
        
        # 转为正式党员日期
        file.active.cell(row=countx, column=22).value = "-"
               
        # 人员类别
        file.active.cell(row=countx, column=23).value = "-"
             
        # 党籍状态
        file.active.cell(row=countx, column=24).value = "-"
            
        # 入党类型
        file.active.cell(row=countx, column=25).value = "-"
        
        # 所在党支部
        file.active.cell(row=countx, column=26).value = "-"
        file.save(path)
        # 人员类型
        file.active.cell(row=countx, column=27).value = "积极分子"
        print("填写第",count,"个积极分子",name_temp,"信息成功") 
        amount_activist_complete = amount_activist_complete + 1



    elif control == 5:
        count = count + 1
        countx = count + 1
        temp_countx = count
        #填写序号
        file.active.cell(row=countx, column=1).value = count
       
        #姓名
        name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
        file.active.cell(row=countx, column=2).value = name_temp
        file.save(path)
        #循环断言
        while(1):
            try:
                df = file.active.cell(row=countx, column=2).value
                assert df != ""
                break
            except AssertionError:
                time.sleep(0.5)
                file.active.cell(row=countx, column=2).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
                file.save(path)
        #性别
        file.active.cell(row=countx, column=3).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[2]//span[1]"))).text
        
        #公民身份证号码
        file.active.cell(row=countx, column=4).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[4]//span[1]"))).text
        
        #民族
        file.active.cell(row=countx, column=5).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[6]//span[1]"))).text
        
        #出生日期
        file.active.cell(row=countx, column=6).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[4]//span[1]"))).text
        
        #学历
        bitian_located(file, path, countx, column=7, wait = wait, str = "(//tbody)[1]/tr[2]/td[6]//span[1]")

        #申请入党日期
        file.active.cell(row=countx, column=8).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[2]//span[1]"))).text
        
        #手机号码
        file.active.cell(row=countx, column=9).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[4]//span[1]"))).text
        
        #背景信息
        bitian_located(file, path, countx, column=10, wait = wait, str = "(//tbody)[1]/tr[3]/td[6]//span[1]")

        #工作岗位
        bitian_located(file, path, countx, column=11, wait = wait, str = "(//tbody)[1]/tr[4]/td[2]//span[1]")
        
        #政治面貌
        bitian_located(file, path, countx, column=12, wait = wait, str = "(//tbody)[1]/tr[4]/td[4]//span[1]")


        #接受申请党组织
        file.active.cell(row=countx, column=13).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[5]/td[2]//span[1]"))).text
        file.save(path)
        if "入党积极分子基本信息" in driver0.page_source:
            #切换到积极分子信息
            while(1):
                try:
                    company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
                    break
                except:
                    time.sleep(0.1)
            while(1):
                try:
                    company_info.click()
                    break
                except:
                    company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
            #籍贯
            file.active.cell(row=countx, column=14).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[2]/tr[3]/td[4]//span[1]"))).text
            
            #入团日期
            file.active.cell(row=countx, column=15).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[4]//span)[1]"))).text
            
            #参加工作日期
            file.active.cell(row=countx, column=16).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[6]//span)[1]"))).text
            
            #申请入党日期
            file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[2]//span)[1]"))).text
            
            #确定入党积极分子日期
            file.active.cell(row=countx, column=18).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[4]//span)[1]"))).text
            
            #工作单位及职务
            file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[6]/td[4]//span)[1]"))).text
            
            #家庭住址
            file.active.cell(row=countx, column=20).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[7]/td[2]//span)[1]"))).text
            
            # 加入党组织日期
            file.active.cell(row=countx, column=21).value = "-"
            
            # 转为正式党员日期
            file.active.cell(row=countx, column=22).value = "-"
               
            # 人员类别
            file.active.cell(row=countx, column=23).value = "-"
                 
            # 党籍状态
            file.active.cell(row=countx, column=24).value = "-"
                
            # 入党类型
            file.active.cell(row=countx, column=25).value = "-"
           
            # 所在党支部
            file.active.cell(row=countx, column=26).value = "-"
            file.save(path)
            # 人员类型
            file.active.cell(row=countx, column=27).value = "入党申请人"
            file.save(path)
            print("填写第",count,"个入党申请人",name_temp,"信息成功") 
            amount_applicant_complete = amount_applicant_complete + 1
        else:
            #籍贯
            file.active.cell(row=countx, column=14).value = "-"
            
            #入团日期
            file.active.cell(row=countx, column=15).value = "-"
           
            #参加工作日期
            file.active.cell(row=countx, column=16).value = "-"
            
            #申请入党日期
            file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr[3]/td[2]"))).text
           
            #确定入党积极分子日期
            file.active.cell(row=countx, column=18).value = "-"
            
            #工作单位及职务
            file.active.cell(row=countx, column=19).value = "-"
           
            #家庭住址
            file.active.cell(row=countx, column=20).value = "-"
          
            # 加入党组织日期
            file.active.cell(row=countx, column=21).value = "-"
           
            # 转为正式党员日期
            file.active.cell(row=countx, column=22).value = "-"
                  
            # 人员类别
            file.active.cell(row=countx, column=23).value = "-"
                  
            # 党籍状态
            file.active.cell(row=countx, column=24).value = "-"
                  
            # 入党类型
            file.active.cell(row=countx, column=25).value = "-"
            
            # 所在党支部
            file.active.cell(row=countx, column=26).value = "-"
            file.save(path)
            # 人员类型
            file.active.cell(row=countx, column=27).value = "入党申请人"

            print("填写第",count,"个入党申请人",name_temp,"信息成功") 
            file.save(path)
            amount_applicant_complete = amount_applicant_complete + 1



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



def wait_click_xpath_action(driver, wait, time_w, xpath):
    while(1):
        try:
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            break
        except:
            time.sleep(time_w)
    while(1):
        try:
            ActionChains(driver).move_to_element(button).click().perform()
            break
        except Exception as e:
            print(f"An exception occurred: {e}")
            button = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            time.sleep(time_w)
