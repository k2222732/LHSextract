import dev_func
from selenium.webdriver.support.ui import WebDriverWait
import login


def main():
    driver = login._main()
    wait = WebDriverWait(driver, 10, 0.5)
    dev_func.access_dev_database(driver, wait)
    dev_func.switch_role(wait, driver)
    #指定位置创建excel工作簿
    dev_func.new_excel(wait)
    dev_func.stop_event.clear()


if __name__ == "__main__":
    main()