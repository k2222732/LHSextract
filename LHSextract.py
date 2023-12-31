import booting
from selenium.webdriver.support.ui import WebDriverWait
chrome_path = "C:/Users/Administrator/AppData/Roaming/360se6/Application/360se.exe"
chromedriver_path = r'd:/project/LHSextract/package/chromedriver.exe'
account = "370830198309261711"
password = "Kfq123456"
url = "http://10.242.32.4:7122/sso/login"

def main():
    global member_excel
    global chrome_path
    global chromedriver_path
    global account
    global password
    global url
    driver = booting.driver_create(chrome_path, chromedriver_path)
    wait = WebDriverWait(driver, 10, 0.5)
    member_excel = 0
    member_total_amount = 0
    booting.login(account, password, driver, url, wait)
    booting.access_member_database(driver, wait)
    booting.switch_role(wait)
    #获取党员总人数
    member_total_amount = booting.get_amountof_member(wait)
    #设置一次性爬取的条目数
    booting.set_amount_perpage(wait)
    #指定位置创建excel工作簿
    booting.new_excel(wait, member_total_amount)

    input("Press Enter to exit...")
if __name__ == "__main__":
    main()