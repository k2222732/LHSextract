import dev_func
from selenium.webdriver.support.ui import WebDriverWait
chrome_path = "C:/Users/Administrator/AppData/Roaming/360se6/Application/360se.exe"
chromedriver_path = r'g:/project/LHSextract/package/chromedriver.exe'
account = "370830198309261711"
password = "Kfq123456"
url = "http://10.242.32.4:7122/sso/login"

def main():
    global chrome_path
    global chromedriver_path
    global account
    global password
    global url
    driver = dev_func.driver_create(chrome_path, chromedriver_path)
    wait = WebDriverWait(driver, 10, 0.5)
    dev_func.login(account, password, driver, url, wait)
    dev_func.access_org_database(driver, wait)
    dev_func.switch_role(wait)
    dev_func.switch_item_org(wait)
    #指定位置创建excel工作簿
    dev_func.new_excel(wait, driver)

    input("Press Enter to exit...")
if __name__ == "__main__":
    main()