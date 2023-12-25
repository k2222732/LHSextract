from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

chrome_options = Options()
chrome_options.binary_location = "C:/Users/Administrator/AppData/Roaming/360se6/Application/360se.exe"  # 替换为你的 Chrome 浏览器的实际安装路径

# ChromeDriver 的路径
chromedriver_path = r'G:/project/LHSextract/package/chromedriver.exe'

# 创建 Service 对象并指定 ChromeDriver 的路径
service = Service(executable_path=chromedriver_path)

# 启动 WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# 设置隐式等待
driver.implicitly_wait(5)

# 打开网址
driver.get("http://10.242.32.4:7122/sso/login")

# 之后可以添加更多的操作，如登录操作等



# 填写登录信息
# 注意：以下ID 'username', 'password', 'captcha', 和 'login_button' 需要根据实际网页元素进行替换
username = driver.find_element(By.ID, 'username')
password = driver.find_element(By.ID, 'password')
captcha = driver.find_element(By.CSS_SELECTOR, '.js-submit.tianze-loginbtn')

username.send_keys("370830198309261711")
password.send_keys("Kfq123456")

# 等待手动输入验证码
print("Please enter the captcha and hit enter in the browser")


# 点击登录按钮
login_button = driver.find_element(By.ID, 'login_button')
login_button.click()

# 之后可以添加额外的代码来处理登录后的页面或关闭浏览器
