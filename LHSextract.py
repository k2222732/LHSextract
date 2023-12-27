from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.binary_location = "C:/Users/Administrator/AppData/Roaming/360se6/Application/360se.exe"  # 替换为你的 Chrome 浏览器的实际安装路径

# ChromeDriver 的路径
chromedriver_path = r'G:/project/LHSextract/package/chromedriver.exe'

# 创建 Service 对象并指定 ChromeDriver 的路径
service = Service(executable_path=chromedriver_path)

# 启动 WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# 设置隐式等待
driver.implicitly_wait(10)

# 打开网址
driver.get("http://10.242.32.4:7122/sso/login")
# 之后可以添加更多的操作，如登录操作等

# 填写登录信息
# 注意：以下ID 'username', 'password', 'captcha', 和 'login_button' 需要根据实际网页元素进行替换
username = driver.find_element(By.ID, 'username')
password = driver.find_element(By.ID, 'password')
validatecode = driver.find_element(By.ID, 'validateCode')
login_button = driver.find_element(By.CSS_SELECTOR, '.js-submit.tianze-loginbtn')

username.send_keys("370830198309261711")
password.send_keys("Kfq123456")

# 等待手动输入验证码
temp = input ("Please enter the captcha and hit enter in the browser")
validatecode.send_keys(temp)
# 点击登录按钮

login_button.click()
# 之后可以添加额外的代码来处理登录后的页面或关闭浏览器

Databaseofparty = driver.find_elements(By.XPATH, '//img[contains(@src, "党组织和党员信息库.png")]')[1]
Databaseofparty.click()


for handle in driver.window_handles:
    driver.switch_to.window(handle)
    if '党组织和党员信息' in driver.title:
        break

droplist = driver.find_element(By.CSS_SELECTOR, '.avatar-wrapper.fs-dropdown-selfdefine')

droplist.click()


wait = WebDriverWait(driver, 10)

role = wait.until(EC.presence_of_element_located((By.XPATH, '//SPAN[contains(text(), "中国共产党山东汶上经济开发区工作委员会-具有审批预备党员权限的基层党委管理员")]')))

role.click()

#在容量较大的磁盘上新建一个文件夹命名为"今天的日期&党员信息库"
#在这个文件夹里新建一个excel文件同样命名为"今天的日期&党员信息库"
#保存"共1155条"到本地变量
#保存"100条/页"到本地变量
#计算总页数到本地变量



input("Press Enter to exit...")
