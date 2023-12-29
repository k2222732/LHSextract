import booting
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

chrome_path = "C:/Users/Administrator/AppData/Roaming/360se6/Application/360se.exe"
chromedriver_path = r'G:/project/LHSextract/package/chromedriver.exe'
account = "370830198309261711"
password = "Kfq123456"
url = "http://10.242.32.4:7122/sso/login"
driver = booting.driver_create(chrome_path, chromedriver_path)
wait = WebDriverWait(driver, 10)


def main():
    booting.login(account, password, driver, url)
    booting.access_member_database(driver)
    booting.switch_role(driver)
    booting.new_excel()
    #保存"共1155条"到本地变量
    #保存"100条/页"到本地变量
    #计算总页数到本地变量
    #

    
    member = driver.find_element(By.XPATH, "(//table[@class='fs-table__body'])[3]/tbody/tr[1]/td[3]")
    member.click()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()