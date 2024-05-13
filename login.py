from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import tkinter as tk
import configparser
from globalv import url
chrome_path = ""
chromedriver_path = ""
account = ""
password = ""

def main():
    global chrome_path
    global chromedriver_path
    global account
    global password
    global url

    config_file = 'config.ini'
    config = configparser.ConfigParser()
    config.read(config_file)
    chrome_path = config.get('Paths_explore', 'explore_path', fallback='')
    chromedriver_path = config.get('Paths_driver', 'explore_driver_path', fallback='')
    account = config.get('Account', 'account', fallback='')
    password = config.get('Password', 'password', fallback='')

    driver = driver_create(chrome_path, chromedriver_path)
    wait = WebDriverWait(driver, 10, 0.5)
    login(account, password, driver, url, wait)
    input("按任意键退出")

def _main():
    global chrome_path
    global chromedriver_path
    global account
    global password
    global url
    config_file = 'config.ini'
    config = configparser.ConfigParser()
    config.read(config_file)
    chrome_path = config.get('Paths_explore', 'explore_path', fallback='')
    chromedriver_path = config.get('Paths_driver', 'explore_driver_path', fallback='')
    account = config.get('Account', 'account', fallback='')
    password = config.get('Password', 'password', fallback='')

    driver = driver_create(chrome_path, chromedriver_path)
    wait = WebDriverWait(driver, 10, 0.5)
    login(account, password, driver, url, wait)
    return driver



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
        get_captcha()
        validatecode.send_keys(Keys.CONTROL + "a")
        validatecode.send_keys(Keys.BACKSPACE)
        validatecode.send_keys(captcha)
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



def get_captcha():
    window = tk.Tk()
    window.title("请输入验证码")
    window.geometry("300x100")

    # 创建标签
    label = tk.Label(window, text="请输入验证码(不区分大小写):")
    label.pack()

    # 创建输入框
    entry = tk.Entry(window)
    entry.pack()

    # 定义获取输入值的函数
    def submit(event = None):
        global captcha
        captcha = entry.get()
        window.destroy()
 
    entry.bind("<Return>", submit)
    # 创建按钮
    button = tk.Button(window, text="确定", command=submit)
    button.pack()

    # 运行窗口
    window.mainloop()



if __name__ == "__main__":
    main()