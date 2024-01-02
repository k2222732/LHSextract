from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from selenium import webdriver
from datetime import datetime
import pandas as pd
import time 
import os



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
    # 打开网址
    driver.get(url)
    username_box = wait.until(EC.visibility_of_element_located((By.ID, 'username')))
    password_box = wait.until(EC.visibility_of_element_located((By.ID, 'password')))
    validatecode = wait.until(EC.visibility_of_element_located((By.ID, 'validateCode')))
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.js-submit.tianze-loginbtn')))
    username_box.send_keys(account)
    password_box.send_keys(password)
    # 等待手动输入验证码
    temp = input ("Please enter the captcha and hit enter in the browser")
    validatecode.send_keys(temp)
    # 点击登录按钮
    login_button.click()
    # 之后可以添加额外的代码来处理登录后的页面或关闭浏览器
    print(f"登录成功")


def access_member_database(driver, wait):
    Databaseofparty = wait.until(EC.element_to_be_clickable((By.XPATH, '(//img[contains(@src, "党组织和党员信息库.png")])[2]')))
    Databaseofparty.click()
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
    time.sleep(3)
    role = wait.until(EC.element_to_be_clickable((By.XPATH, '//SPAN[contains(text(), "中国共产党山东汶上经济开发区工作委员会-具有审批预备党员权限的基层党委管理员")]')))
    role.click()
    print(f"切换角色成功")


def new_excel():
    #在当前文件夹新建一个文件夹命名为"database"在里面新建一个文件夹名为"database_member"
    os.makedirs("database/database_member", exist_ok=True)
    #在database_member文件夹里新建一个excel文件命名为"今天的日期"&"党员信息库"
    today_date = datetime.now().strftime("%Y-%m-%d")
    excel_file_name = f"{today_date}党员信息库.xlsx"
    excel_file_path = os.path.join("database", "database_member", excel_file_name)
    #在excel里的第一行建立表头，分别是"姓名、性别、公民身份号码、民族、出生日期、学历、人员类别、学位、所在党支部、手机号码、入党日期
    #转正日期、党龄、党龄校正值、新社会阶层类型、工作岗位、从事专业技术职务、是否农民工、现居住地、户籍所在地、是否失联党员、是否流动党员、
    #党员注册时间、注册手机号、党员增加信息、党员减少信息、入党类型、转正情况、入党时所在支部、延长预备期时间
    columns = ["序号","姓名", "性别", "公民身份号码", "民族", "出生日期", "学历", "人员类别", "学位", 
           "所在党支部", "手机号码", "入党日期", "转正日期", "党龄", "党龄校正值", "新社会阶层类型", 
           "工作岗位", "从事专业技术职务", "是否农民工", "现居住地", "户籍所在地", "是否失联党员", 
           "是否流动党员", "党员注册时间", "注册手机号", "党员增加信息", "党员减少信息", "入党类型", 
           "转正情况", "入党时所在支部", "延长预备期时间"]
    
    df = pd.DataFrame(columns=columns)
    df.to_excel(excel_file_path, index=False)

    workbook = load_workbook(excel_file_path)
    print(f"文件 '{excel_file_path}' 已成功创建。")
    return workbook, excel_file_path


def cycle(wait):
    member = wait.until(EC.element_to_be_clickable((By.XPATH, "(//table[@class='fs-table__body'])[3]/tbody/tr[1]/td[3]")))
    member.click() 
    print(f"进入党员个人页面成功")   

def synchronizing(count, file, wait, path):
    count = count +1
    countx = count + 1
    #在member_excel的第countx行第一列存countx
    file.active.cell(row=countx, column=1).value = count
    file.save(path)
    temp_name = wait.until(EC.visibility_of_element_located((By.XPATH, "(//div[@class = 'card-class']//div[@class = 'fs-tabs__content']//div[@class = 'row-val-shot'])[1]")))
    temp = temp_name.text
    file.active.cell(row=countx, column=2).value = temp
    file.save(path)
    print(f"填写党员姓名成功")  


