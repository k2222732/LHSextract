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


amount_that_complete = 0
from globalv import org_directory


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
    while(1):
        try:
            Databaseofparty = wait.until(EC.element_to_be_clickable((By.XPATH, '(//img[contains(@src, "党组织和党员信息库.png")])[2]')))
            break
        except:
            time.sleep(0.5)
    while(1):
        try:
            Databaseofparty.click()
            break
        except:

            Databaseofparty = wait.until(EC.element_to_be_clickable((By.XPATH, '(//img[contains(@src, "党组织和党员信息库.png")])[2]')))
            time.sleep(0.5)

    try:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if '党组织和党员信息' in driver.title:
                break
        print(f"进入党组织和党员信息库成功")
    except:
        time.sleep(1)


def switch_role(wait):
    droplist = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.avatar-wrapper.fs-dropdown-selfdefine')))
    droplist.click()
    print(f"进入角色下拉列表成功")
    time.sleep(1)
    role = wait.until(EC.element_to_be_clickable((By.XPATH, '//SPAN[contains(text(), "中国共产党山东汶上经济开发区工作委员会-具有审批预备党员权限的基层党委管理员")]')))
    role.click()
    print(f"切换角色成功")



#切换到党组织信息页面
def switch_item_org(wait):
    while(1):
        try:
            item_org = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[contains(text(), "党组织管理")]/..')))
            break
        except:
            time.sleep(0.1)
    while(1):
        try:
            item_org.click()
            break
        except:
            item_org = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[contains(text(), "党组织管理")]/..')))
    while(1):
        try:
            org_info = wait.until(EC.element_to_be_clickable((By.XPATH, '(//*[contains(text(), "信息管理")]/..)[2]')))
            break
        except:
            time.sleep(0.1)
    while(1):
        try:
            org_info.click()
            break
        except:
            org_info = wait.until(EC.element_to_be_clickable((By.XPATH, '(//*[contains(text(), "信息管理")]/..)[2]')))


def new_excel(wait, driver):
    global org_directory
    today_date = datetime.now().strftime("%Y-%m-%d")
    excel_file_name = f"{today_date}党组织信息库.xlsx"
    excel_file_path = os.path.join(org_directory, excel_file_name)
    if os.path.isfile(excel_file_path):
        #org_excel = load_workbook(excel_file_path)
        #rebuild(excel_file_path, wait, org_excel, excel_file_path)


        #创建目录g:/project/LHSextract/database/database_org
        os.makedirs(org_directory, exist_ok=True)
        #设计党组织基本信息表头
        columns_base_info = ["序号","党组织全称", "组织树", "党组织简称", "党内统计用党组织简称", "成立日期", "党组织编码", "党组织联系人", "联系电话", "组织类别", 
            '是否具有"审批预备党员权限"', "功能型党组织", "党组织所在单位情况", "党组织所在行政区划", "批准成立的上级党组织", "是否为新业态", "驻外情况", 
            "党组织曾用名", "单位名称（全称）", "UUID", "有无统一社会信用代码", "法人单位统一社会信用代码", "单位性质类别", 
            "法人单位标识", "建立党组情况", "法人单位建立党组织情况", "在岗职工人数", "企业控制（控股）情况", "企业规模", "单位所在目录", "单位隶属关系", "单位所在行政区划", "单位名称(全称)", "机构类型", "法人单位统一社会信用代码"
            , "新经济行业", "经济行业", "经济类型", "新经济类型", "成立日期", "注册地行政区划", "注册地址", "组织机构代码", "上级主管部门名称", "单位隶属关系"]


        #创建一个dataframe表头为columns_base_info中的元素
        df = pd.DataFrame(columns = columns_base_info)
        #dataframe导出到excel
        df.to_excel(excel_file_path, index = False)
        #使用openpyxl库来加载一个已存在的Excel工作簿
        org_excel = load_workbook(excel_file_path)
        #打印调试信息
        print(f"文件 '{excel_file_path}' 已成功创建。")
        #启动同步
        synchronizing(wait, org_excel, excel_file_path, driver)
    else:
        #创建目录g:/project/LHSextract/database/database_org
        os.makedirs(org_directory, exist_ok=True)
        #设计党组织基本信息表头
        columns_base_info = ["序号","党组织全称", "组织树", "党组织简称", "党内统计用党组织简称", "成立日期", "党组织编码", "党组织联系人", "联系电话", "组织类别", 
            '是否具有"审批预备党员权限"', "功能型党组织", "党组织所在单位情况", "党组织所在行政区划", "批准成立的上级党组织", "是否为新业态", "驻外情况", 
            "党组织曾用名", "单位名称（全称）", "UUID", "有无统一社会信用代码", "法人单位统一社会信用代码", "单位性质类别", 
            "法人单位标识", "建立党组情况", "法人单位建立党组织情况", "在岗职工人数", "企业控制（控股）情况", "企业规模", "单位所在目录", "单位隶属关系", "单位所在行政区划", "单位名称(全称)", "机构类型", "法人单位统一社会信用代码"
            , "新经济行业", "经济行业", "经济类型", "新经济类型", "成立日期", "注册地行政区划", "注册地址", "组织机构代码", "上级主管部门名称", "单位隶属关系"]


        #创建一个dataframe表头为columns_base_info中的元素
        df = pd.DataFrame(columns = columns_base_info)
        #dataframe导出到excel
        df.to_excel(excel_file_path, index = False)
        #使用openpyxl库来加载一个已存在的Excel工作簿
        org_excel = load_workbook(excel_file_path)
        #打印调试信息
        print(f"文件 '{excel_file_path}' 已成功创建。")
        #启动同步
        synchronizing(wait, org_excel, excel_file_path, driver)
  


def access_info_page(wait, row):
    xpath = f"(//table[@class='fs-table__body'])[3]/tbody/tr[{row}]/td[3]"
    while 1:
        try:
            org = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            break
        except:
            time.sleep(0.1)
    while 1:
        try:
            org.click()
            break
        except:
                while 1:
                    try:
                        org = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                        break
                    except:
                        time.sleep(0.1)
    print(f"进入党员个人页面成功") 


def rebuild(excel_file_path, wait, org_total_amount, org_excel, org_excel_path):
    #先改变全局变量
    init_complete_amount(excel_file_path)
    synchronizing(wait, org_excel, org_excel_path)


def init_complete_amount(excel_file_path):
    global amount_that_complete
    df = pd.read_excel(excel_file_path, sheet_name=0)
    row_count = df.dropna(how='all').shape[0]
    amount_that_complete = row_count - 1


def synchronizing(wait, org_excel, org_excel_path, driver):
    #递归遍历树状列表完成每一个党组织的信息采集。
    #点击根组织(//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@class = "fs-tree-node__content"])[1]
    while (1):
        try:
            root_org = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class = 'tree_wrapper']//div[@role = 'treeitem']/div[@class = 'fs-tree-node__content'])[1]")))
            break
        except:
            time.sleep(0.5)
    while(1):
        try:
            root_org.click()
            break
        except:
            root_org = wait.until(EC.element_to_be_clickable((By.XPATH, "(//div[@class = 'tree_wrapper']//div[@role = 'treeitem']/div[@class = 'fs-tree-node__content'])[1]")))
            time.sleep(0.5)
    #采集根组织信息
    downloading(file = org_excel, wait = wait, driver = driver, path = org_excel_path)

    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    while(1):
        try:
            tree_root = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@class = 'tree_wrapper']//div[@role = 'treeitem']/div[@role = 'group']")))
            break
        except:
            time.sleep(0.5)


    item_html = tree_root.get_attribute('outerHTML')
    with open('item_html.txt', 'w', encoding = 'utf-8') as file0:
        file0.write(item_html)

    #recursion(tree_root)
    recursion(tree_root = tree_root, file = org_excel, wait = wait, driver = driver, path = org_excel_path)


def recursion(tree_root, file, wait, driver, path):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    
    #item_html0 = org_items[0].get_attribute('outerHTML')
    #with open('item_html0.txt', 'w', encoding = 'utf-8') as file0:
        #file0.write(item_html0)

    #for 每个元素 in tree_root
    for item in org_items:
        # 每个元素.click()
        item.click()
        # 采集党组织信息
        downloading(file, wait, driver, path)
        # 查看元素下面是否有//span[@class = "is-leaf fs-tree-node__expand-icon fs-icon-caret-right"]

        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        first_child = soup.find()
        first_grandchild = first_child.find() if first_child else None
        leaf = first_grandchild.find('span', class_= "is-leaf fs-tree-node__expand-icon fs-icon-caret-right")
        #如果有：
        if leaf:
            #next
            next
        #如果没有：
        else:
            #断言元素下面有//span[@class = "fs-tree-node__expand-icon fs-icon-caret-right"]
                #断言异常
            #点击//span[@class = "fs-tree-node__expand-icon fs-icon-caret-right"]
            while (1):
                try:
                    arrow_down = item.find_element(By.XPATH, "./div[@class='fs-tree-node__content']/span[@class='fs-tree-node__expand-icon fs-icon-caret-right']")
                    break
                except:
                    time.sleep(0.5)
            while(1):
                try:
                    arrow_down.click()
                    break
                except:
                    time.sleep(0.5)
                    arrow_down = item.find_element(By.XPATH, "./div[@class='fs-tree-node__content']/span[@class='fs-tree-node__expand-icon fs-icon-caret-right']")
            #元素下面的//div[@role = "group"]结构体作为新的根节点结构体new_tree_root
            while(1):
                try:
                    new_tree_root = item.find_element(By.XPATH, ".//div[@role = 'group']")
                    break
                except:
                    time.sleep(0.5)
            #递归recursion(new_tree_root)
            recursion(new_tree_root, file, wait, driver, path)

def downloading(file, wait, driver, path):
    global amount_that_complete
    count = amount_that_complete
    count = count + 1
    countx = count + 1
    #填写序号#
    file.active.cell(row=countx, column=1).value = count
    file.save(path)
    
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
            name_temp = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), '党组织全称')]/../following-sibling::*/div[1]"))).text
            file.active.cell(row=countx, column=2).value = name_temp
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



def table_council(org_name: str, wait, driver):
    #获取数据存储目录
    global org_directory
    #获取当天日期
    today_date = datetime.now().strftime("%Y-%m-%d")
    #构建文件名称
    excel_file_name = f"{today_date}{org_name}班子成员信息库.xlsx"
    #构建完整路径
    excel_file_path = os.path.join(org_directory, excel_file_name)
    #创建一个dataframe表头为columns_base_info中的元素
    df = pd.DataFrame()
    #dataframe导出到excel
    df.to_excel(excel_file_path, index = False)
    #openpyexcel打开excel_file_path
    book = load_workbook(excel_file_path)
    sheet = book.active
    #将党组织的全称存放在单元格A1
    sheet['A1'] = org_name
    book.save(excel_file_path)
    #设计党组织委员会信息表头
    table_head = ["序号", "党内职务", "姓名", "公民身份证号码", "性别", "出生日期", "学历", "领导职务", "任职日期", "离职日期", "排序", "公司职务"]
    for i, value in enumerate(table_head, start = 1):
        sheet.cell(row = 2, column = i, value = value)
        book.save(excel_file_path)
    #读取表格容器保存在变量container里
    while(1):
        try:
            tbody = wait.until(EC.visibility_of_element_located((By.XPATH, "(//span[contains(text(), '班子信息集')]/../../../div[3]//tbody)[1]")))
                                                                           #(//div[@class = 'fs-table__fixed-body-wrapper'])[3]//tbody
                                                                           #//div[@class = 'fs-table fs-table--fit fs-table--border fs-table--enable-row-transition fs-table--mini']/div[@class = 'fs-table__body-wrapper is-scrolling-none']//tbody
                                                                           #(//div[@class = 'fs-table__body-wrapper is-scrolling-left'])[5]//tbody
            break
        except:
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
            erro_html = driver.execute_script("return document.documentElement.outerHTML")
            with open('erro.txt', 'w', encoding = "utf-8") as file1:
                file1.write(erro_html)
            break

    #用beautifulsoup分析container，将表头保存在新的容器container_header里，将表体保存在新的容器container_body里
   # while(1):
        #try:
            #soup = BeautifulSoup(tbody, 'html.parser')
            #break
        #except TypeError:
            #time.sleep(0.5)
            #tbody = wait.until(EC.visibility_of_element_located((By.XPATH, "(//div[@class = 'fs-table__header-wrapper']/following-sibling::div[@class = 'fs-table__body-wrapper is-scrolling-left'])[2]//tbody")))
    soup = BeautifulSoup(tbody.get_attribute("outerHTML"), 'html.parser')
    tr = soup.find_all('tr')
    #从tr里提取数据保存在自单元格A3始向右的区域内
    for j, every_row in enumerate(tr, start = 1):

        # 从elm_tr里解析出td, 保存在td数组里
        col = every_row.find_all('td')
        
        #将序号i填入第j+2行，第1列
        sheet.cell(row = j+2, column = 1, value = j)
        book.save(excel_file_path)


        #将党内职务填入第j+2行，第2列
        position = col[1].find_all('span')[-1]
        position_content = position.get_text(strip = True)
        sheet.cell(row = j+2, column = 2, value = position_content)
        book.save(excel_file_path)


        #将姓名填入第j+2行，第3列
        name = col[2].find_all('span')[-1]
        name_content = name.get_text(strip = True)
        sheet.cell(row = j+2, column = 3, value = name_content)
        book.save(excel_file_path)


        #将公民身份证号码职务填入第j+2行，第4列
        idcard_no = col[3].find_all('div')[-1]
        idcard_no_content = idcard_no.get_text(strip = True)
        sheet.cell(row = j+2, column = 4, value = idcard_no_content)
        book.save(excel_file_path)


        #将性别填入第j+2行，第5列
        gender = col[4].find_all('span')[-1]
        gender_content = gender.get_text(strip = True)
        sheet.cell(row = j+2, column = 5, value = gender_content)
        book.save(excel_file_path)


        #将出生日期填入第j+2行，第6列
        birthday = col[5].find_all('span')[-1]
        birthday_content = birthday.get_text(strip = True)
        sheet.cell(row = j+2, column = 6, value = birthday_content)
        book.save(excel_file_path)


        #将学历填入第j+2行，第7列
        edu_qual = col[6].find_all('span')[-1]
        edu_qual_content = edu_qual.get_text(strip = True)
        sheet.cell(row = j+2, column = 7, value = edu_qual_content)
        book.save(excel_file_path)


        #将领导职务填入第j+2行，第8列span
        leader_position = col[7].find_all('span')[-1]
        leader_position_content = leader_position.get_text(strip = True)
        sheet.cell(row = j+2, column = 8, value = leader_position_content)
        book.save(excel_file_path)


        #将任职日期填入第j+2行，第9列span
        appointmentdate = col[8].find_all('span')[-1]
        appointmentdate_content = appointmentdate.get_text(strip = True)
        sheet.cell(row = j+2, column = 9, value = appointmentdate_content)
        book.save(excel_file_path)


        #将离职日期填入第j+2行，第10列span
        resignationdate = col[9].find_all('span')[-1]
        resignationdate_content = resignationdate.get_text(strip = True)
        sheet.cell(row = j+2, column = 10, value = resignationdate_content)
        book.save(excel_file_path)


        #将排序填入第j+2行，第11列div
        sort = col[10].find_all('div')[-1]
        sort_content = sort.get_text(strip = True)
        sheet.cell(row = j+2, column = 11, value = sort_content)
        book.save(excel_file_path)


        #将公司职务填入第j+2行，第12列div
        companyposition = col[11].find_all('div')[-1]
        companyposition_content = companyposition.get_text(strip = True)
        sheet.cell(row = j+2, column = 12, value = companyposition_content)
        book.save(excel_file_path)

    #释放资源
    book.close()
    


def table_reward_punish(org_name: str, wait):
    #获取数据存储目录
    global org_directory
    #获取当天日期
    today_date = datetime.now().strftime("%Y-%m-%d")
    #构建文件名称
    excel_file_name = f"{today_date}{org_name}奖惩信息.xlsx"
    #构建完整路径
    excel_file_path = os.path.join(org_directory, excel_file_name)
    #创建一个dataframe表头为columns_base_info中的元素
    df = pd.DataFrame()
    #dataframe导出到excel
    df.to_excel(excel_file_path, index = False)
    #openpyexcel打开excel_file_path
    book = load_workbook(excel_file_path)
    sheet = book.active
    #将党组织的全称存放在单元格A1
    sheet['A1'] = org_name
    book.save(excel_file_path)
    #设计党组织委员会信息表头
    table_head = ["序号", "奖惩名称", "批准机关", "批准日期"]
    #将设计好的表头填入excel
    for i, value in enumerate(table_head, start = 1):
        sheet.cell(row = 2, column = i, value = value)
        book.save(excel_file_path)
    #寻找奖惩块
    while 1: 
        try:
            box_table = wait.until(EC.visibility_of_element_located((By.ID, "pane-rewardsPunishments")))
            break        
        except:
            time.sleep(0.5)
    #soup分析奖惩块
    table_html = box_table.get_attribute('outerHTML')
    soup = BeautifulSoup(table_html, 'html.parser')
    #检查是否有数据，如果没有则全填-
    if soup.find_all(string = lambda text: '暂无数据' in text):
        for cell in sheet['A3:D3'][0]:
            cell.value = '-'
            book.save(excel_file_path)
    #如果有数据则填表
    else:
        #寻找表体tbody
        while 1:
            try:
                tbody = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@class = 'fs-table__body-wrapper is-scrolling-none']//tbody")))
                break
            except:
                time.sleep(0.5)
        #soup分析tbody
        tbody_html = tbody.get_attribute("outerHTML")
        soup = BeautifulSoup(tbody_html, 'html.parser')
        #soup提取所有行
        rows= soup.find_all('tr')
        
        #对于每一行
        for j, each_row in enumerate(rows, start = 1):
            #soup提取所有列
            col = each_row.find_all('td')

            #将序号j填入第j+2行，第1列
            sheet.cell(row = j+2, column = 1, value = j)
            book.save(excel_file_path)

            #将奖惩名称填入第j+2行，第2列
            try:
                name = col[0].find_all('div')[-1]
            except:
                name = col[0].find_all('span')[-1]
            name_content = name.get_text(strip = True)
            sheet.cell(row = j+2, column = 2, value = name_content)
            book.save(excel_file_path)

            #将批准机关填入第j+2行，第3列
            try:
                justice = col[1].find_all('span')[-1]
            except:
                justice = col[1].find_all('div')[-1]
            justice_content = justice.get_text(strip = True)
            sheet.cell(row = j+2, column = 3, value = justice_content)
            book.save(excel_file_path)

            #将批准日期填入第j+2行，第4列
            try:
                passdate = col[2].find_all('span')[-1]
            except:
                passdate = col[2].find_all('div')[-1]
            passdate_content = passdate.get_text(strip = True)
            sheet.cell(row = j+2, column = 4, value = passdate_content)
            book.save(excel_file_path)

    #释放资源
    book.close()


def element_exists(driver, by, value):
    try:
        driver.find_element(by, value)
        return True
    except NoSuchElementException:
        return False
    




