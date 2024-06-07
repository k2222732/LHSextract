import dev_func
from base.waitclick import wait_click_xpath
from selenium.webdriver.support.ui import WebDriverWait
import login
import base.write_entry as input
import base.uploadpic as uppic
import os
import pandas as pd
from tkinter import messagebox
import base.rwconfig as conf

applicant_info = {}

def main():
    driver = login._main()
    wait = WebDriverWait(driver, 10, 0.5)
    dev_func.access_dev_database(driver, wait)
    dev_func.switch_role(wait, driver)
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "发展流程")])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "添加")]')
    mfilepath = conf.pick_config_param(configfilepath = '..\\config.ini', str_class = '', str_param = '')
    global applicant_info
    applicant_info = load_to_cache(mfilepath)
    if applicant_info == False:
        return False
    step_1(driver, wait)
    step_2(driver, wait)
    step_3(driver, wait)
    step_4(driver, wait)
    step_5(driver, wait)
    

#从模板中读取积极分子的基本信息保存到字典中
def load_to_cache(mfilepath):
    '''
    mfilepath: 采集信息的xls文件夹路径
    返回值：字典
    '''
    files = [f for f in os.listdir(mfilepath) if f.endswith('.xlsx') and f == '信息采集表.xlsx']
    if len(files) != 1:
            messagebox.showinfo("提示","目录中必须包含且仅包含一个 '信息采集表.xlsx' 文件")
            return False
    file_path = os.path.join(mfilepath, files[0])
    df = pd.read_excel(file_path)
    data_dict = {}
    for index, row in df.iterrows():#df.iterrows()方法返回一个迭代器,该迭代器会生成DataFrame中每行的索引和行数据。具体来说,每次迭代时,它会返回一个包含两个元素的元组
            name = row['Name']
            age = row['Age']
            role = row['Role']
            data_dict[name] = {'Age': age, 'Role': role}
    return data_dict



def step_1(driver, wait):
    #import base.write_entry as input
    #import base.uploadpic as uppic
    pass

def step_2(driver, wait):
    pass

def step_3(driver, wait):
    pass

def step_4(driver, wait):
    pass

def step_5(driver, wait):
    pass





if __name__ == "__main__":
    main()