import dev_func
from selenium.webdriver.support.ui import WebDriverWait
from globalv import chrome_path, chromedriver_path, chromedriver_path, account, password, url


def main():
    global chrome_path
    global chromedriver_path
    global account
    global password
    global url
    driver = dev_func.driver_create(chrome_path, chromedriver_path)
    wait = WebDriverWait(driver, 10, 0.5)
    dev_func.login(account, password, driver, url, wait)
    dev_func.access_dev_database(driver, wait)
    dev_func.switch_role(wait, driver)
    #指定位置创建excel工作簿
    dev_func.new_excel(wait)
    dev_func.stop_event.clear()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()