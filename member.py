import member_func
from selenium.webdriver.support.ui import WebDriverWait
from globalv import chrome_path, chromedriver_path, account, password, url

def main():
    global member_excel
    global chrome_path
    global chromedriver_path
    global account
    global password
    global url
    driver = member_func.driver_create(chrome_path, chromedriver_path)
    wait = WebDriverWait(driver, 10, 0.5)
    member_excel = 0
    member_total_amount = 0
    member_func.login(account, password, driver, url, wait)
    member_func.access_member_database(driver, wait)
    member_func.switch_role(wait)
    #获取党员总人数
    member_total_amount = member_func.get_amountof_member(wait)
    #设置一次性爬取的条目数
    member_func.set_amount_perpage(wait)
    #指定位置创建excel工作簿
    member_func.new_excel(wait, member_total_amount)

    input("Press Enter to exit...")
if __name__ == "__main__":
    main()