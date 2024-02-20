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


directory = "g:/project/LHSextract/database/database_dev"

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


#切换到正式党员信息页面
def switch_formal_mem(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[2]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]/li[@class = "el-select-dropdown__item hover"]/span[contains(text(), "正式党员")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')
    
def switch_informal_mem(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[2]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]/li[@class = "el-select-dropdown__item"]/span[contains(text(), "预备党员")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')

def switch_devtarg(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[2]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]/li[@class = "el-select-dropdown__item"]/span[contains(text(), "发展对象")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')

def switch_activist(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[2]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]/li[@class = "el-select-dropdown__item"]/span[contains(text(), "入党积极分子")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')

def switch_applicant(wait):
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)[2]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]/li[@class = "el-select-dropdown__item"]/span[contains(text(), "入党申请人")]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '//span[contains(text(), "搜索")]')

def new_excel(wait):
    global directory
    today_date = datetime.now().strftime("%Y-%m-%d")
    excel_file_name = f"{today_date}党员发展纪实信息库.xlsx"
    excel_file_path = os.path.join(directory, excel_file_name)
    #if os.path.isfile(excel_file_path):
        #member_excel = load_workbook(excel_file_path)
        #rebuild(excel_file_path, wait, member_total_amount, member_excel, excel_file_path)
    #else:
        #在当前文件夹新建一个文件夹命名为"database"在里面新建一个文件夹名为"database_member"
        #os.makedirs(directory, exist_ok=True)
        #在database_member文件夹里新建一个excel文件命名为"今天的日期"&"党员信息库"

        #在excel里的第一行建立表头，分别是"姓名、性别、公民身份号码、民族、出生日期、学历、人员类别、学位、所在党支部、手机号码、入党日期
        #转正日期、党龄、党龄校正值、新社会阶层类型、工作岗位、从事专业技术职务、是否农民工、现居住地、户籍所在地、是否失联党员、是否流动党员、
        #党员注册时间、注册手机号、党员增加信息、党员减少信息、入党类型、转正情况、入党时所在支部、延长预备期时间
    columns = ["姓名","公民身份证号码","民族","性别","出生日期","学历","申请入党日期","手机号码","背景信息","工作岗位",
               "政治面貌","接受申请党组织","籍贯","入团日期","参加工作日期","申请入党日期","确定入党积极分子日期",
               "工作单位及职务","家庭住址","加入党组织日期","转为正式党员日期","人员类别","党籍状态","入党类型","所在党支部"]
        
    df = pd.DataFrame(columns=columns)
    df.to_excel(excel_file_path, index=False)
    member_excel = load_workbook(excel_file_path)
    print(f"文件 '{excel_file_path}' 已成功创建。")
    switch_formal_mem(wait)
    synchronizing(wait, member_excel, excel_file_path, control=1)
    switch_informal_mem(wait)
    synchronizing(wait, member_excel, excel_file_path, control=2)
    switch_devtarg(wait)
    synchronizing(wait, member_excel, excel_file_path, control=3)
    switch_activist(wait)
    synchronizing(wait, member_excel, excel_file_path, control=4)
    switch_applicant(wait)
    synchronizing(wait, member_excel, excel_file_path, control=5)

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


def schedule(complete, total, xpath, wait, member_excel, member_excel_path, control):
    global amount_mem_complete
    global amount_infomem_complete
    global amount_activist_complete
    global amount_devtar_complete
    global amount_applicant_complete
    while complete < total:
        page_number = int(complete / 100 + 1)
        row_number = int(complete % 100 + 1)
        #time.sleep(0.1)
        input_page = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        input_page.send_keys(Keys.CONTROL + "a")
        input_page.send_keys(Keys.BACKSPACE)
        input_page.send_keys(page_number)
        input_page.send_keys(Keys.RETURN)
        access_info_page(wait, row_number)
        while 1:
            try:
                downloading(complete, member_excel, wait, member_excel_path, control)
                break
            except:
                time.sleep(0.5)
                access_info_page(wait, row_number)
    if control == 1:
        amount_mem_complete = amount_mem_complete + 1
    elif control == 2:
        amount_infomem_complete = amount_infomem_complete + 1
    elif control == 3:
        amount_devtar_complete = amount_devtar_complete + 1
    elif control == 4:
        amount_activist_complete = amount_activist_complete + 1
    elif control == 5:
        amount_applicant_complete = amount_applicant_complete + 1



def access_info_page(wait, row):
    wait_click_xpath(wait, time_w = 0.5, xpath = f"(//table[@class='fs-table__body'])[3]/tbody/tr[{row}]/td[3]")
    print(f"进入党员个人页面成功") 


def downloading(file, wait, driver, path, control):
    global amount_mem_complete
    global amount_infomem_complete
    global amount_activist_complete
    global amount_devtar_complete
    global amount_applicant_complete
    if control == 1:
        count = amount_mem_complete
    elif control == 2:
        count = amount_infomem_complete
    elif control == 3:
        count = amount_activist_complete
    elif control == 4:
        count = amount_devtar_complete
    elif control == 5:
        count = amount_applicant_complete
    
    count = count + 1
    countx = count + 1
    #填写序号#
    file.active.cell(row=countx, column=1).value = count
    file.save(path)
    ##########################################################################
    #填写党组织全称#
    name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '党组织全称')]/../following-sibling::*/div[1]"))).text
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
            file.active.cell(row=countx, column=2).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '党组织全称')]/../following-sibling::*/div[1]"))).text
            file.save(path)
    #组织树
    file.active.cell(row=countx, column=3).value = "-" #wait.until(EC.presence_of_element_located((By.XPATH, "(//div[@class = 'card-class']//div[@class = 'fs-tabs__content']//div[@class = 'row-val-shot'])[3]"))).text
    file.save(path)
    #党组织简称#
    file.active.cell(row=countx, column=4).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '党组织简称')]/../following-sibling::*/div[1]"))).text
    file.save(path)
    #党内统计用党组织简称#
    file.active.cell(row=countx, column=5).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),  '党内统计用党组织简称')]/following-sibling::*/div[1]"))).text
    file.save(path)
    #成立日期#
    file.active.cell(row=countx, column=6).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),  '成立日期')]/../following-sibling::div/div[1]"))).text
    file.save(path)
    #党组织编码#
    file.active.cell(row=countx, column=7).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),  '党组织编码')]/following-sibling::*/span"))).text
    file.save(path)
    #党组织联系人#
    file.active.cell(row=countx, column=8).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),  '党组织联系人')]/following-sibling::*/div[1]"))).text
    file.save(path)
    #联系电话#
    file.active.cell(row=countx, column=9).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '联系电话')]/../following-sibling::*/div[1]"))).text
    file.save(path)
    #组织类别#
    org_type0 = file.active.cell(row=countx, column=10).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '组织类别')]/../following-sibling::*/div[1]"))).text
    file.save(path)
    #是否具有"审批预备党员权限"#
    if "委员会" in org_type0:
        file.active.cell(row=countx, column=11).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),  '审批预备党员权限')]/following-sibling::div/span"))).text
        file.save(path)
    elif "总支" in org_type0:
        file.active.cell(row=countx, column=11).value = "-"
        file.save(path)
    else:
        file.active.cell(row=countx, column=11).value = "-"
        file.save(path)
    #功能型党组织#
    file.active.cell(row=countx, column=12).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),  '功能型党组织')]/following-sibling::*/span"))).text
    file.save(path)
    #党组织所在单位情况#
    file.active.cell(row=countx, column=13).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '党组织所在单位情况')]/../following-sibling::*/div[1]"))).text
    file.save(path)
    #党组织所在行政区划#
    file.active.cell(row=countx, column=14).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '党组织所在行政区划')]/../following-sibling::*/div[1]"))).text
    file.save(path)
    #批准成立的上级党组织#
    file.active.cell(row=countx, column=15).value = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),  '批准成立的上级党组织')]/../following-sibling::*/div[1]"))).text
    file.save(path)
    #是否为新业态#
    file.active.cell(row=countx, column=16).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),  '是否为新业态')]/following-sibling::*/div[1]"))).text
    file.save(path)
    #驻外情况#
    file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(),  '驻外情况')]/following-sibling::*/div[1]"))).text
    file.save(path)
    #党组织曾用名，这里需要用soup判断是否有这一条
    html0 = driver.find_element(By.XPATH, '//div[contains(text(), "党组织曾用名")]/../..')
    _html0 = html0.get_attribute("style")
    if "display: none;" in _html0:
    #党组织曾用名，这里需要用soup判断是否有这一条
        file.active.cell(row=countx, column=18).value = "无"
        file.save(path)
    else:
        file.active.cell(row=countx, column=18).value = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(),  '党组织曾用名')]/following-sibling::*//tbody/tr/td[2]/div/div"))).text
        file.save(path)


    #切换选项卡
    while(1):
        try:
            company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '党组织单位信息')]")))
            break
        except:
            time.sleep(0.1)
    while(1):
        try:
            company_info.click()
            break
        except:
            company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '党组织单位信息')]")))


    #单位名称（全称）#


    file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '单位名称（全称）')]/following-sibling::div"))).text
    file.save(path)
    i_1 = 0
    while(i_1 < 15):
        try:
            df = file.active.cell(row=countx, column=19).value
            file.save(path)
            print("df:",df)
            assert df != ""
            break
        except AssertionError:
            time.sleep(0.5)
            file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '单位名称（全称）')]/following-sibling::div"))).text
            file.save(path)
            i_1 = i_1 + 1


    #UUID#
    file.active.cell(row=countx, column=20).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), 'UUID')]/following-sibling::div"))).text
    file.save(path)
    #有无统一社会信用代码#
    file.active.cell(row=countx, column=21).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '有无统一社会信用代码')]/following-sibling::div"))).text
    file.save(path)
    #法人单位统一社会信用代码#
    file.active.cell(row=countx, column=22).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '法人单位统一社会信用代码')]/following-sibling::div"))).text
    file.save(path)
    #单位性质类别#
    org_type = file.active.cell(row=countx, column=23).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '单位性质类别')]/following-sibling::div"))).text
    file.save(path)
    #法人单位标识#
    file.active.cell(row=countx, column=24).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '法人单位标识')]/following-sibling::div"))).text
    file.save(path)
    #建立党组情况!
    temp_1 = wait.until(EC.presence_of_element_located((By.XPATH, "(//div[@class = 'info-content-view'])[2]")))
    temp_1_html = temp_1.get_attribute("outerHTML")
    if '建立党组情况' in temp_1_html:
        file.active.cell(row=countx, column=25).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '建立党组情况')]/following-sibling::div"))).text
        file.save(path)
    else:
        file.active.cell(row=countx, column=25).value = "-"
        file.save(path)
    #法人单位建立党组织情况!
    if '法人单位建立党组织情况' in temp_1_html:
        file.active.cell(row=countx, column=26).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '法人单位建立党组织情况')]/following-sibling::div"))).text
        file.save(path)
    else:
        file.active.cell(row=countx, column=26).value = "-"
        file.save(path)
    #在岗职工数#
    #如果字符串org_type里面包含字眼"公司"则执行下面两行代码
    if "公司" in org_type:
        file.active.cell(row=countx, column=27).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '在岗职工数')]/following-sibling::div"))).text
        file.save(path)
    #否则执行
    else:
        file.active.cell(row=countx, column=27).value = "-"
        file.save(path)
    #在企业控制（控股）情况
    if "公司" in org_type:
        file.active.cell(row=countx, column=28).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '企业控制（控股）情况')]/following-sibling::div"))).text
        file.save(path)
    else:
        file.active.cell(row=countx, column=28).value = "-"
        file.save(path)
    #企业规模
    if "公司" in org_type:
        file.active.cell(row=countx, column=29).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '企业规模')]/following-sibling::div"))).text
        file.save(path)
    else:
        file.active.cell(row=countx, column=29).value = "-"
        file.save(path)
    #单位所在目录
    file.active.cell(row=countx, column=30).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '单位所在目录')]/following-sibling::div"))).text
    file.save(path)
    #如果单位性质类别位行政机关，则补充单位隶属关系
    if "机关" in org_type:
        file.active.cell(row=countx, column=45).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '隶属关系')]/following-sibling::div/div/span"))).text
        file.save(path)
    else:
        file.active.cell(row=countx, column=45).value = "-"
        file.save(path)

    #民营科技企业标识
    if "公司" in org_type:
        file.active.cell(row=countx, column=31).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '民营科技企业标识')]/following-sibling::div"))).text
        file.save(path)
    else:
        file.active.cell(row=countx, column=31).value = "-"
        file.save(path)
    # 单位所在行政区划
    file.active.cell(row=countx, column=32).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '单位所在行政区划')]/following-sibling::div//input"))).text
    file.save(path)
    # 判断参考信息（省标院等单位）是否存在
    while 1:
        try:
            readonly = wait.until(EC.presence_of_element_located((By.XPATH, "(//div[@class = 'read-only'])[1]")))
            break
        except:
            time.sleep(0.5)
    
    readonly_html = readonly.get_attribute('outerHTML')
    soup = BeautifulSoup(readonly_html, 'html.parser')
    
    if soup.find_all(string= lambda text: '机构类型' in text):
        # 单位名称(全称)
        file.active.cell(row=countx, column=33).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '单位名称(全称)')]/following-sibling::div"))).text
        file.save(path)
        # 机构类型
        file.active.cell(row=countx, column=34).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '机构类型')]/following-sibling::div//span[@style = 'margin-top: auto; margin-bottom: auto;']"))).text
        file.save(path)        
        # 法人单位统一社会信用代码
        file.active.cell(row=countx, column=35).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), '法人单位统一社会信用代码')])[2]/following-sibling::div"))).text
        file.save(path)        
        # 新经济行业
        file.active.cell(row=countx, column=36).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '新经济行业')]/following-sibling::div"))).text
        file.save(path)       
        # 经济行业
        file.active.cell(row=countx, column=37).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), '经济行业')]/following-sibling::div)[2]"))).text
        file.save(path)
        # 经济类型
        file.active.cell(row=countx, column=38).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//label[contains(text(), '经济类型')]/following-sibling::div)[1]//span[@style = 'margin-top: auto; margin-bottom: auto;']"))).text
        file.save(path)
        # 新经济类型
        file.active.cell(row=countx, column=39).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '新经济类型')]/following-sibling::div//span[@style = 'margin-top: auto; margin-bottom: auto;']"))).text
        file.save(path)
        # 成立日期
        file.active.cell(row=countx, column=40).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '成立日期')]/following-sibling::div"))).text
        file.save(path)
        # 注册地行政区划
        file.active.cell(row=countx, column=41).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '注册地行政区划')]/following-sibling::div"))).text
        file.save(path)
        # 注册地址
        file.active.cell(row=countx, column=42).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '注册地址')]/following-sibling::div"))).text
        file.save(path)
        # 组织机构代码
        file.active.cell(row=countx, column=43).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '组织机构代码')]/following-sibling::div"))).text
        file.save(path)
        # 上级主管部门名称
        file.active.cell(row=countx, column=44).value = wait.until(EC.presence_of_element_located((By.XPATH, "//label[contains(text(), '上级主管部门名称')]/following-sibling::div"))).text
        file.save(path)


    else:
        file.active.cell(row=countx, column=33).value = "-"
        file.save(path)
        # 机构类型
        file.active.cell(row=countx, column=34).value = "-"
        file.save(path)        
        # 法人单位统一社会信用代码
        file.active.cell(row=countx, column=35).value = "-"
        file.save(path)        
        # 新经济行业
        file.active.cell(row=countx, column=36).value = "-"
        file.save(path)       
        # 经济行业
        file.active.cell(row=countx, column=37).value = "-"
        file.save(path)
        # 经济类型
        file.active.cell(row=countx, column=38).value = "-"
        file.save(path)
        # 新经济类型
        file.active.cell(row=countx, column=39).value = "-"
        file.save(path)
        # 成立日期
        file.active.cell(row=countx, column=40).value = "-"
        file.save(path)
        # 注册地行政区划
        file.active.cell(row=countx, column=41).value = "-"
        file.save(path)
        # 注册地址
        file.active.cell(row=countx, column=42).value = "-"
        file.save(path)
        # 组织机构代码
        file.active.cell(row=countx, column=43).value = "-"
        file.save(path)
        # 上级主管部门名称
        file.active.cell(row=countx, column=44).value = "-"
        file.save(path)
    # 切换选项卡到班子成员
    while 1:
        try:
            councilcard = wait.until(EC.element_to_be_clickable((By.ID, "tab-class")))
            break
        except:
            time.sleep(0.5)
    while 1:
        try:
            councilcard.click()
            break
        except:
            councilcard = wait.until(EC.element_to_be_clickable((By.ID, "tab-class")))
            time.sleep(0.5)
    # 采集班子成员信息
    table_council(name_temp, wait, driver)
    # 切换选项卡到惩戒信息
    while 1:
        try:
            RewaridsAndPunishments = wait.until(EC.element_to_be_clickable((By.ID, "tab-rewardsPunishments")))
            break
        except:
            time.sleep(0.5)
    while 1:
        try:
            RewaridsAndPunishments.click()
            break
        except:
            RewaridsAndPunishments = wait.until(EC.element_to_be_clickable((By.ID, "tab-rewardsPunishments")))
            time.sleep(0.5)
    # 采集奖励惩戒信息
    table_reward_punish(name_temp, wait)
    
    
    print("填写第",count,"个党组织",name_temp,"信息成功") 
    amount_that_complete = amount_that_complete + 1






def init_complete_amount(excel_file_path):
    global amount_that_complete
    df = pd.read_excel(excel_file_path, sheet_name=0)
    row_count = df.dropna(how='all').shape[0]
    amount_that_complete = row_count - 1


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


