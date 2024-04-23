import org_func
from selenium.webdriver.support.ui import WebDriverWait
from globalv import url
import configparser
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


    driver = org_func.driver_create(chrome_path, chromedriver_path)
    wait = WebDriverWait(driver, 10, 0.5)
    org_func.login(account, password, driver, url, wait)
    org_func.access_org_database(driver, wait)
    org_func.switch_role(wait)
    org_func.switch_item_org(wait)
    #指定位置创建excel工作簿
    org_func.new_excel(wait, driver)
    org_func.stop_event.clear()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()