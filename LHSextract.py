import booting
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

chrome_path = "C:/Users/Administrator/AppData/Roaming/360se6/Application/360se.exe"
chromedriver_path = r'G:/project/LHSextract/package/chromedriver.exe'
account = "370830198309261711"
password = "Kfq123456"
url = "http://10.242.32.4:7122/sso/login"

def main():
    driver = booting.driver_create(chrome_path, chromedriver_path)
    booting.login(account, password, driver, url)
    booting.access_member_database(driver)
    booting.switch_role(driver)

    #在容量较大的磁盘上新建一个文件夹命名为"今天的日期&党员信息库"
    #在这个文件夹里新建一个excel文件同样命名为"今天的日期&党员信息库"
    #在excel里的第一行建立表头，分别是"姓名、性别、公民身份号码、民族、出生日期、学历、人员类别、学位、所在党支部、手机号码、入党日期
    #转正日期、党龄、党龄校正值、新社会阶层类型、工作岗位、从事专业技术职务、是否农民工、现居住地、户籍所在地、是否失联党员、是否流动党员、党员
    #党员注册时间、注册手机号、党员增加信息、党员减少信息、入党类型、转正情况、入党时所在支部、延长预备期时间
    #保存"共1155条"到本地变量
    #保存"100条/页"到本地变量
    #计算总页数到本地变量
    #

    
    member = driver.find_element(By.XPATH, "(//table[@class='fs-table__body'])[3]/tbody/tr[1]/td[3]")
    member.click()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()