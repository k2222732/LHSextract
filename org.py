import org_func
from selenium.webdriver.support.ui import WebDriverWait
import login
from app.merge_xls import *


def main():
    driver = login._main()
    wait = WebDriverWait(driver, 10, 0.5)
    
    org_func.access_org_database(driver, wait)
    org_func.switch_role(wait)
    org_func.switch_item_org(wait)
    #指定位置创建excel工作簿
    org_func.new_excel(wait, driver)
    org_func.stop_event.clear()
    merge()


if __name__ == "__main__":
    main()