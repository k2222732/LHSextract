import dev_func
from base.waitclick import *
from selenium.webdriver.support.ui import WebDriverWait
import login
import os
import pandas as pd
from tkinter import messagebox
import base.rwconfig as conf
import base.write_entry as cursor
from base.boot import synchronizing_org, synchronizing_job, synchronizing_xueli, synchronizing_jg
from base.mystruct import TreeNode
import time
from base.solv_date import *
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import dev_func
from bs4 import BeautifulSoup
from base.write_entry import *
from selenium.webdriver.common.keys import Keys


applicant_info = {}


def main():
    result_report = []
    org_tree = TreeNode()
    job_tree = TreeNode()
    education_tree = TreeNode()
    jg = TreeNode()
    driver = login._main()
    wait = WebDriverWait(driver, 2, 0.5)
    dev_func.access_dev_database(driver, wait)
    dev_func.switch_role(wait, driver)
    #载入党员信息表
    #mfilepath = conf.pick_config_param(configfilepath = '..\\config.ini', str_class = '', str_param = '')
    mfilepath = './/材料图片保存处'
    global applicant_info
    applicant_info = load_to_cache(mfilepath)
    if applicant_info == False:
        return False
    
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "发展流程")])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = "//div[@class = 'sub-step-container']//span[contains(text(),'递交入党申请书')]")
    wait_click_xpath(wait, time_w = 0.5, xpath = "//span[contains(text(), '添加入党申请人')]")
    #同步党组织树、职务树、学历树
    wait_click_xpath(wait, time_w=0.5, xpath="//button[@class = 'el-button fr fas el-button--default el-button--small']")
    synchronizing_org(driver=driver, wait=wait, input_node=org_tree)
    #org_tree.dayin()
    synchronizing_job(driver=driver, wait=wait, input_node=job_tree)
    #job_tree.dayin()
    synchronizing_xueli(driver=driver, wait=wait, input_node=education_tree)

    #循环录入积极分子信息
    person_old = ""
    for person in applicant_info:
        if person == person_old:
            print("循环失败" )
            raise Exception("循环失败")
        card1 =None
        card2 =None
        card3 =None
        card4 =None
        card5 =None
        card6 =None
        card7 =None
        wait_click_xpath(wait, time_w = 0.5, xpath = '(//*[contains(text(), "人员信息")]/..)')
        cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
        time.sleep(2)
        cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
        cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入全量公民身份号码']", text=applicant_info[person]['公民身份证号码'])
        time.sleep(2)
        cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入全量公民身份号码']", text=applicant_info[person]['公民身份证号码'])
        wait_click_xpath(wait, time_w = 0.5, xpath = '(//input[@placeholder = "请选择"])[1]')
        wait_click_xpath(wait, time_w = 0.5, xpath = '//ul[@class = "el-scrollbar__view el-select-dropdown__list"]//li//span[contains(text(), "全部")]')
        wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "查询")])[3]')
        #number = dev_func.get_total_amount(wait, "//span[@class = 'el-pagination__total']")
        t = wait_return_subelement_absolute_notmust(wait, time_w= 0.5, xpath="//span[contains(text(), '暂无数据')]", times=3)
        if t == None:
            number = 1
        else :
            number = 0
        
        if number == 0:
            wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "发展流程")])[1]')
            wait_click_xpath(wait, time_w = 0.5, xpath = "//div[@class = 'sub-step-container']//span[contains(text(),'递交入党申请书')]")
            wait_click_xpath(wait, time_w = 0.5, xpath = "//span[contains(text(), '添加入党申请人')]")
            
        
            
            #education_tree.dayin()
            step_1(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], org_tree =org_tree, job_tree=job_tree, education_tree=education_tree)
            step_2(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
            step_3(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], jg=jg)
            step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
            step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
            result_report.append(f"积极分子{applicant_info[person]['姓名']} {applicant_info[person]['公民身份证号码']}")
        else:
            #看这个人是如当申请人还是积极分子  //table[@class = 'el-table__body']//tr//td[8]
            rylx = wait_return_subelement_absolute(wait, time_w=0.5, xpath="//table[@class = 'el-table__body']//tr//td[9]") #2024/8/21/11/07改
            rylx = rylx.get_attribute('textContent')

            #切换到发展流程
            ryxx = wait_return_subelement_absolute(wait, time_w=0.5, xpath="//span[contains(text(), '发展流程')]")
            ryxx.click()
            #选择申请人或积极分子
            if rylx == "入党申请人 ":
                wait_click_xpath(wait, time_w=0.5, xpath="(//div[contains(text(), '申请入党')])[1]")
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
                time.sleep(2)
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入公民身份号码']", text=applicant_info[person]['公民身份证号码'])
                time.sleep(2)
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入公民身份号码']", text=applicant_info[person]['公民身份证号码'])
                wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "查询")])[3]')
                try:
                    click_button_until_specifyxpath_appear_except(wait, driver, specifyxpath = "//span[contains(text(), '模板下载')]", buttonxpath = "//tbody[1]/tr[1]/td[2]//span", time_w=5, times=3)
                except TimeoutException:
                    wait_click_xpath(wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子')])[1]")
                    cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
                    time.sleep(2)
                    cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
                    cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入公民身份号码']", text=applicant_info[person]['公民身份证号码'])
                    time.sleep(2)
                    cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入公民身份号码']", text=applicant_info[person]['公民身份证号码'])
                    wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "查询")])[3]')
                    wait_click_xpath(wait, time_w=0.5, xpath="//tbody[1]/tr[1]/td[2]//span")
                    jijifenzi(wait, driver, org_tree, job_tree, education_tree, person, jg)
                    continue

                #click_button_until_specifyxpath_appear(wait, driver, "//span[contains(text(), '模板下载')]", "//tbody[1]/tr[1]/td[2]//span")
            #输入身份证号查询
            #点击此人进入发展流程
            #获得‘提交入党申请书’元素，检查css (//div[contains(text(), '递交入党申请书')])[10]

                card1a = wait_return_element_two_xpath(driver, xpath1= "(//div[contains(text(), '递交入党申请书')])[31]", xpath2= "(//div[contains(text(), '递交入党申请书')])[2]", timewait = 1)
                card2a = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '党组织派人谈话')])", times=1) 
                card3a = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '推荐和确定入党积极分子')])", times=1) 
                card4a = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子公示和备案')])", times=1) 
                card5a = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '指定培养联系人')])", times=1) 
                card6a = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子培养教育考察')])", times=1) 
                card7a = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '确定发展对象人选')])", times=1)


                card1b = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '递交入党申请书')])", times=1) 
                card2b = wait_return_element_two_xpath(driver, xpath1= "(//div[contains(text(), '党组织派人谈话')])[31]", xpath2= "(//div[contains(text(), '党组织派人谈话')])[2]", timewait = 1)
                card3b = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '推荐和确定入党积极分子')])", times=1) 
                card4b = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子公示和备案')])", times=1) 
                card5b = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '指定培养联系人')])", times=1) 
                card6b = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子培养教育考察')])", times=1) 
                card7b = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '确定发展对象人选')])", times=1) 
                if all(value is not None and value is not False for value in [card1a, card2a, card3a, card4a, card5a, card6a, card7a]):
                    card1 = card1a
                    card2 = card2a
                    card3 = card3a
                    card4 = card4a
                    card5 = card5a
                    card6 = card6b
                    card7 = card7a
                elif all(value is not None and value is not False for value in [card1b, card2b, card3b, card4b, card5b, card6b, card7b]):
                    card1 = card1b
                    card2 = card2b
                    card3 = card3b
                    card4 = card4b
                    card5 = card5b
                    card6 = card6b
                    card7 = card7b







                


                card1_html = card1.get_attribute('outerHTML')
                soup1 = BeautifulSoup(card1_html, 'html.parser')
                element1 = soup1.find(lambda tag: tag.name == "div" and "递交入党申请书" in tag.text)
                style1 = element1.get('style')

                card2_html = card2.get_attribute('outerHTML')
                soup2 = BeautifulSoup(card2_html, 'html.parser')
                element2 = soup2.find(lambda tag: tag.name == "div" and "党组织派人谈话" in tag.text)
                style2 = element2.get('style')

                card3_html = card3.get_attribute('outerHTML')
                soup3 = BeautifulSoup(card3_html, 'html.parser')
                element3 = soup3.find(lambda tag: tag.name == "div" and "推荐和确定入党积极分子" in tag.text)
                style3 = element3.get('style')

                card4_html = card4.get_attribute('outerHTML')
                soup4 = BeautifulSoup(card4_html, 'html.parser')
                element4 = soup4.find(lambda tag: tag.name == "div" and "入党积极分子公示和备案" in tag.text)
                style4 = element4.get('style')

                card5_html = card5.get_attribute('outerHTML')
                soup5 = BeautifulSoup(card5_html, 'html.parser')
                element5 = soup5.find(lambda tag: tag.name == "div" and "指定培养联系人" in tag.text)
                style5 = element5.get('style')

                card6_html = card6.get_attribute('outerHTML')
                soup6 = BeautifulSoup(card6_html, 'html.parser')
                element6 = soup6.find(lambda tag: tag.name == "div" and "入党积极分子培养教育考察" in tag.text)
                style6 = element6.get('style')

                card7_html = card7.get_attribute('outerHTML')
                soup7 = BeautifulSoup(card7_html, 'html.parser')
                element7 = soup7.find(lambda tag: tag.name == "div" and "确定发展对象人选" in tag.text)
                style7 = element7.get('style')
                

                if 'cursor: pointer;' in style1 and "cursor: not-allowed;" in style2:
                    card1.click()
                    step_1_rebuild(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], org_tree =org_tree, job_tree=job_tree, education_tree=education_tree)
                    step_2(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                    step_3(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], jg=jg)
                    step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                    step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                elif 'cursor: pointer;' in style2 and "cursor: not-allowed;" in style3:
                    card2.click()
                    step_2(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                    step_3(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], jg=jg)
                    step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                    step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                elif 'cursor: pointer;' in style3 and "cursor: not-allowed;" in style4:
                    card3.click()
                    step_3(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], jg=jg)
                    step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                    step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                elif 'cursor: pointer;' in style4 and "cursor: not-allowed;" in style5:
                    card4.click()
                    step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                    step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                elif 'cursor: pointer;' in style5 and "cursor: not-allowed;" in style6:
                    card5.click()
                    step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
                elif 'cursor: pointer;' in style6 and "cursor: not-allowed;" in style7:
                    pass
                
                
             
             #判断执行
            #获得‘党组织派人谈话’元素，检查css
             #判断执行
            #获得‘推荐和确定入党积极分子’元素，检查css
             #判断执行
            #获得‘入党积极分子公示和备案’元素，检查css
             #判断执行
            #获得‘制定培养联系人’元素，检查css
             #判断执行  
            elif rylx == "入党积极分子 ":
                wait_click_xpath(wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子')])[1]")
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
                time.sleep(2)
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入姓名']", text=applicant_info[person]['姓名'])
                
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入公民身份号码']", text=applicant_info[person]['公民身份证号码'])
                time.sleep(2)
                cursor.input_text(wait, driver, xpath="//input[@placeholder = '请输入公民身份号码']", text=applicant_info[person]['公民身份证号码'])
                wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "查询")])[3]')
                
                
                click_button_until_specifyxpath_appear_except(wait, driver, specifyxpath = "//span[contains(text(), '模板下载')]", buttonxpath = "//tbody[1]/tr[1]/td[2]//span", time_w=5, times=5)
                #click_button_until_specifyxpath_appear(wait, driver, "//span[contains(text(), '模板下载')]", "//tbody[1]/tr[1]/td[2]//span")
            #输入身份证号查询
            #点击此人进入发展流程
            #获得‘提交入党申请书’元素，检查css (//div[contains(text(), '递交入党申请书')])[10]
                jijifenzi(wait, driver, org_tree, job_tree, education_tree, person, jg)
    person_old = person
    input("按任意键结束")



def jijifenzi(wait, driver, org_tree, job_tree, education_tree, person, jg):
    
    card1c = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=1, xpath="(//div[contains(text(), '递交入党申请书')])", times=3) 
    card2c = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '党组织派人谈话')])", times=1) 
    card3c = wait_return_element_two_xpath(driver, xpath1= "(//div[contains(text(), '推荐和确定入党积极分子')])[31]", xpath2= "(//div[contains(text(), '推荐和确定入党积极分子')])[2]", timewait = 1)
    card4c = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子公示和备案')])", times=1) 
    card5c = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '指定培养联系人')])", times=1) 
    card6c = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子培养教育考察')])", times=1) 
    card7c = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '确定发展对象人选')])", times=1)


    card1d = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '递交入党申请书')])", times=1) 
    card2d = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '党组织派人谈话')])", times=1) 
    card3d = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '推荐和确定入党积极分子')])", times=1) 
    card4d = wait_return_element_two_xpath(driver, xpath1= "(//div[contains(text(), '入党积极分子公示和备案')])[1]", xpath2= "(//div[contains(text(), '入党积极分子公示和备案')])[31]", timewait = 1)
    card5d = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '指定培养联系人')])", times=1) 
    card6d = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子培养教育考察')])", times=1) 
    card7d = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '确定发展对象人选')])", times=1) 


    card1e = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '递交入党申请书')])", times=1) 
    card2e = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '党组织派人谈话')])", times=1) 
    card3e = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '推荐和确定入党积极分子')])", times=1) 
    card4e = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子公示和备案')])", times=1) 
    card5e = wait_return_element_two_xpath(driver, xpath1= "(//div[contains(text(), '指定培养联系人')])[31]", xpath2= "(//div[contains(text(), '指定培养联系人')])[2]", timewait = 1)
    card6e = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子培养教育考察')])", times=1) 
    card7e = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '确定发展对象人选')])", times=1) 

    
    card1f = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '递交入党申请书')])", times=1) 
    card2f = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '党组织派人谈话')])", times=1) 
    card3f = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '推荐和确定入党积极分子')])", times=1) 
    card4f = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子公示和备案')])", times=1) 
    card5f = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '指定培养联系人')])", times=1) 
    card6f = wait_return_element_two_xpath(driver, xpath1= "(//div[contains(text(), '入党积极分子培养教育考察')])[31]", xpath2= "(//div[contains(text(), '入党积极分子培养教育考察')])[2]", timewait = 1)
    card7f = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '确定发展对象人选')])", times=1) 



    card1g = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '递交入党申请书')])", times=1) 
    card2g = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '党组织派人谈话')])", times=1) 
    card3g = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '推荐和确定入党积极分子')])", times=1) 
    card4g = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子公示和备案')])", times=1) 
    card5g = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '指定培养联系人')])", times=1) 
    card6g = wait_return_subelement_absolute_notmust_v1(driver, wait, time_w=0.5, xpath="(//div[contains(text(), '入党积极分子培养教育考察')])", times=1) 
    card7g = wait_return_element_two_xpath(driver, xpath1= "(//div[contains(text(), '确定发展对象人选')])[31]", xpath2= "(//div[contains(text(), '确定发展对象人选')])[2]", timewait = 1)



    if all(value is not None and value is not False for value in [card1c, card2c, card3c, card4c, card5c, card6c, card7c]):
        card1 = card1c
        card2 = card2c
        card3 = card3c
        card4 = card4c
        card5 = card5c
        card6 = card6c
        card7 = card7c
    elif all(value is not None and value is not False for value in [card1d, card2d, card3d, card4d, card5d, card6d, card7d]):
        card1 = card1d
        card2 = card2d
        card3 = card3d
        card4 = card4e
        card5 = card5d
        card6 = card6d
        card7 = card7d
    elif all(value is not None and value is not False for value in [card1e, card2e, card3e, card4e, card5e, card6e, card7e]):
        card1 = card1e
        card2 = card2e
        card3 = card3e
        card4 = card4e
        card5 = card5e
        card6 = card6e
        card7 = card7e
    elif all(value is not None and value is not False for value in [card1f, card2f, card3f, card4f, card5f, card6f, card7f]):
        card1 = card1f
        card2 = card2f
        card3 = card3f
        card4 = card4f
        card5 = card5f
        card6 = card6f
        card7 = card7f


    

    
    card1_html = card1.get_attribute('outerHTML')
    soup1 = BeautifulSoup(card1_html, 'html.parser')
    element1 = soup1.find(lambda tag: tag.name == "div" and "递交入党申请书" in tag.text)
    style1 = element1.get('style')

    card2_html = card2.get_attribute('outerHTML')
    soup2 = BeautifulSoup(card2_html, 'html.parser')
    element2 = soup2.find(lambda tag: tag.name == "div" and "党组织派人谈话" in tag.text)
    style2 = element2.get('style')

    card3_html = card3.get_attribute('outerHTML')
    soup3 = BeautifulSoup(card3_html, 'html.parser')
    element3 = soup3.find(lambda tag: tag.name == "div" and "推荐和确定入党积极分子" in tag.text)
    style3 = element3.get('style')

    card4_html = card4.get_attribute('outerHTML')
    soup4 = BeautifulSoup(card4_html, 'html.parser')
    element4 = soup4.find(lambda tag: tag.name == "div" and "入党积极分子公示和备案" in tag.text)
    style4 = element4.get('style')

    card5_html = card5.get_attribute('outerHTML')
    soup5 = BeautifulSoup(card5_html, 'html.parser')
    element5 = soup5.find(lambda tag: tag.name == "div" and "指定培养联系人" in tag.text)
    style5 = element5.get('style')

    card6_html = card6.get_attribute('outerHTML')
    soup6 = BeautifulSoup(card6_html, 'html.parser')
    element6 = soup6.find(lambda tag: tag.name == "div" and "入党积极分子培养教育考察" in tag.text)
    style6 = element6.get('style')


    card7_html = card7.get_attribute('outerHTML')
    soup7 = BeautifulSoup(card7_html, 'html.parser')
    element7 = soup7.find(lambda tag: tag.name == "div" and "确定发展对象人选" in tag.text)
    style7 = element7.get('style')


    if 'cursor: pointer;' in style1 and "cursor: not-allowed;" in style2:
        card1.click()
        step_1_rebuild(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], org_tree =org_tree, job_tree=job_tree, education_tree=education_tree)
        step_2(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
        step_3(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], jg=jg)
        step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
        step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
    elif 'cursor: pointer;' in style2 and "cursor: not-allowed;" in style3:
        card2.click()
        step_2(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
        step_3_rebuild(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], jg=jg)
        step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
        step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
    elif 'cursor: pointer;' in style3 and "cursor: not-allowed;" in style4:
        card3.click()
        step_3_rebuild(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'], jg=jg)
        step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
        step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
    elif 'cursor: pointer;' in style4 and "cursor: not-allowed;" in style5:
        card4.click()
        step_4(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
        step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
    elif 'cursor: pointer;' in style5 and "cursor: not-allowed;" in style6:
        card5.click()
        step_5(driver, wait, data_dict=applicant_info, gmsfzhm=applicant_info[person]['公民身份证号码'])
    elif 'cursor: pointer;' in style6 and "cursor: not-allowed;" in style7:
        pass
    card1 =None
    card2 =None
    card3 =None
    card4 =None
    card5 =None
    card6 =None
    card7 =None
#输入身份证号查询
#点击此人进入发展流程
#获得‘提交入党申请书’元素，检查css
    
    #判断执行
#获得‘党组织派人谈话’元素，检查css
    #判断执行
#获得‘推荐和确定入党积极分子’元素，检查css
    #判断执行
#获得‘入党积极分子公示和备案’元素，检查css
    #判断执行
#获得‘制定培养联系人’元素，检查css
    #判断执行







#从模板中读取积极分子的基本信息保存到字典中
def load_to_cache(mfilepath):
    '''
    mfilepath: 采集信息的xls文件夹路径
    返回值：字典
    '''
    #目标目录下只允许有一个'信息采集表.xlsx'
    files = [f for f in os.listdir(mfilepath) if f.endswith('.xlsx') and f == '信息采集表.xlsx']
    if len(files) != 1:
            messagebox.showinfo("提示","目录中必须包含且仅包含一个 '信息采集表.xlsx' 文件")
            return False
    #通过链接得到绝对路径
    file_path = os.path.join(mfilepath, files[0])

    df = pd.read_excel(file_path)
    data_dict = {}
    for index, row in df.iterrows():#df.iterrows()方法返回一个迭代器,该迭代器会生成DataFrame中每行的索引和行数据。具体来说,每次迭代时,它会返回一个包含两个元素的元组
            xh = row['序号']
            xm = row['姓名']
            xb = row['性别']
            gmsfzhm = row['公民身份证号码']
            mz = row['民族']
            csny = row['出生日期']
            xl = row['学历']
            sqrdrq = row['申请入党日期']
            sjhm = row['手机号码']
            gzgw = row['工作岗位']
            zzmm = row['政治面貌']
            jssqdzz = row['接受申请党组织']
            jg = row['籍贯']
            rtrq = row['入团日期']
            cjgzrq = row['参加工作日期']
            sqrdrq = row['申请入党日期']
            jjfzhyrq = row['积极分子会议日期']
            qdrdjjfzrq = row['确定入党积极分子日期']
            gzdwjzw = row['工作单位及职务']
            jtzz = row['家庭住址']
            pylxr = row['培养联系人']
            qdjjfzdhyzcr = row['确定积极分子的会议主持人']
            dzbsj = row['党支部书记']
            jjfzhydzzy = row['积极分子会议的列席组织员']
            yxqk = row['一线情况']
            jjfzqdhysdrs = row['积极分子确定会议实到人数']
            jjfzqdhyydrs = row['积极分子确定会议应到人数']
            nested_dict = {'序号': xh, '姓名': xm, '性别':xb, '公民身份证号码':gmsfzhm,'民族':mz,
                                  '出生日期':csny, '学历':xl, '申请入党日期':sqrdrq, '手机号码':sjhm, 
                                  '工作岗位':gzgw, '政治面貌':zzmm, '接受申请党组织':jssqdzz, '籍贯':jg ,
                                  '入团日期':rtrq, '参加工作日期':cjgzrq,'积极分子会议日期':jjfzhyrq, '申请入党日期':sqrdrq, '确定入党积极分子日期':qdrdjjfzrq,
                                  '工作单位及职务':gzdwjzw, '家庭住址':jtzz, '培养联系人':pylxr, '确定积极分子的会议主持人':qdjjfzdhyzcr,'党支部书记':dzbsj,
                                  '积极分子会议的列席组织员':jjfzhydzzy, '一线情况': yxqk, '积极分子确定会议实到人数':jjfzqdhysdrs, '积极分子确定会议应到人数':jjfzqdhyydrs}
            data_dict[gmsfzhm]=nested_dict
    return data_dict



def step_1(driver, wait, data_dict, gmsfzhm, org_tree, job_tree, education_tree):
    


    #import base.write_entry as input
    #编辑#(//span[contains(text(), '编辑')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[1]")
    #姓名#(//input[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[1]", text=data_dict[gmsfzhm]['姓名'])
    #公民身份证号码#(//input[@placeholder = '请输入内容'])[2]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[2]", text=data_dict[gmsfzhm]['公民身份证号码'])
    #实名认证#(//span[contains(text(), '实名认证')])[3]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '实名认证')])[3]")
    #民族#(//input[@placeholder = '请选择字段项'])[1]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[1]")
    cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value=data_dict[gmsfzhm]['民族'])
    #性别#(//input[@placeholder = '请选择字段项'])[2]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[2]")
    cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value=data_dict[gmsfzhm]['性别'])
    #出生日期#(//input[@placeholder = '选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[1]", text=data_dict[gmsfzhm]['出生日期'])
    #学历#(//input[@placeholder = '请选择字段项'])[3]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[3]")
    cursor.select_party_org(wait, driver, tree=education_tree, target=data_dict[gmsfzhm]['学历'], element_xpath="//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    #递交入党申请书日期#(//input[@placeholder = '选择日期'])[2]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[2]", text=data_dict[gmsfzhm]['申请入党日期'])
    t = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "(//input[@placeholder = '选择日期'])[2]")
    t.send_keys(Keys.RETURN)
    #手机号码#(//input[@placeholder = '请输入联系方式'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入联系方式'])[1]", text=data_dict[gmsfzhm]['手机号码'])
    #家庭住址#(//input[@placeholder = '请输入内容'])[2]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入家庭住址'])[1]", text=data_dict[gmsfzhm]['家庭住址'])
    #工作单位及职务#(//input[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入工作单位及职务'])[1]", text=data_dict[gmsfzhm]['工作单位及职务'])
    
    
    #背景信息#(//input[@placeholder = '请选择字段项'])[4]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[4]")
    cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value='正常')
    #工作岗位#(//input[@placeholder = '请选择字段项'])[5]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[5]")
    cursor.select_party_org(wait, driver, tree=job_tree, target=data_dict[gmsfzhm]['工作岗位'], element_xpath="//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    #政治面貌#(//input[@placeholder = '请选择字段项'])[6]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[6]")
    cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value=data_dict[gmsfzhm]['政治面貌'])
    #重新入党#(//input[@placeholder = '请选择'])
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择'])")
    cursor.select_background(wait, driver, element_xpath="//ul[@class = 'el-scrollbar__view el-select-dropdown__list']", label='span', value='否')
    #所属党组织#(//input[@placeholder = '请选择字段项'])[7]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[8]")
    cursor.select_party_org(wait, driver, tree=org_tree, target = data_dict[gmsfzhm]['接受申请党组织'], element_xpath="//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    #接受申请党组织#(//input[@placeholder = '请输入接受申请党组织'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入接受申请党组织'])[1]", text=data_dict[gmsfzhm]['接受申请党组织'])
    #一线情况#
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[7]")
    cursor.select_background(wait, driver, "//div[@class = 'el-tree objectTree']", label='span', value = data_dict[gmsfzhm]['一线情况'])
    #保存#(//span[contains(text(), '保存')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '保存')])[1]")
    #编辑#(//span[contains(text(), '编辑')])[2]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[2]")
    #上传图片#//input[@type="file"]
    #需运行前验证文件夹是否存在
    current_working_dir = os.getcwd()
    absolute_path = current_working_dir + f"\\材料图片保存处\\{data_dict[gmsfzhm]['姓名']}+{gmsfzhm}\\入党申请书照片"
    cursor.upload_pic(pic_dir=absolute_path, input_xpath="//input[@type='file']", wait=wait, driver=driver)
    #保存#(//span[contains(text(), '保存')])[2]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="//span[contains(text(), '取消')]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w=0.1, times=2)

    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(),'提交')])[2]", buttonxpath="(//span[contains(text(),'提交')])[2]", time_w=0.1, times=2)

    
    #警告#//div[@role = 'alert']
    #警告的内容#//div[@role = 'alert']//p[@class = 'el-message__content']
   

def step_1_rebuild(driver, wait, data_dict, gmsfzhm, org_tree, job_tree, education_tree):
    #import base.write_entry as input
    #编辑#(//span[contains(text(), '编辑')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[1]")
    #姓名#(//input[@placeholder = '请输入内容'])[1]
    #cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[1]", text=data_dict[gmsfzhm]['姓名'])
    #公民身份证号码#(//input[@placeholder = '请输入内容'])[2]
    #cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[2]", text=data_dict[gmsfzhm]['公民身份证号码'])
    #实名认证#(//span[contains(text(), '实名认证')])[3]
    #cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '实名认证')])[3]")
    #民族#(//input[@placeholder = '请选择字段项'])[1]
    

    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[2]")
    cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value=data_dict[gmsfzhm]['民族'])
    #性别#(//input[@placeholder = '请选择字段项'])[2]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[3]")
    cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value=data_dict[gmsfzhm]['性别'])
    #出生日期#(//input[@placeholder = '选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[1]", text=data_dict[gmsfzhm]['出生日期'])
    #学历#(//input[@placeholder = '请选择字段项'])[3]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[4]")
    cursor.select_party_org(wait, driver, tree=education_tree, target=data_dict[gmsfzhm]['学历'], element_xpath="//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    #递交入党申请书日期#(//input[@placeholder = '选择日期'])[2]
    #cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[2]", text=data_dict[gmsfzhm]['申请入党日期'])
    #手机号码#(//input[@placeholder = '请输入联系方式'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入联系方式'])[1]", text=data_dict[gmsfzhm]['手机号码'])
    #家庭住址#(//input[@placeholder = '请输入内容'])[2]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入家庭住址'])[1]", text=data_dict[gmsfzhm]['家庭住址'])
    #工作单位及职务#(//input[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入工作单位及职务'])[1]", text=data_dict[gmsfzhm]['工作单位及职务'])
    
    
    #背景信息#(//input[@placeholder = '请选择字段项'])[4]
    #cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[4]")
    #cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value='正常')
    #工作岗位#(//input[@placeholder = '请选择字段项'])[5]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[6]")
    cursor.select_party_org(wait, driver, tree=job_tree, target=data_dict[gmsfzhm]['工作岗位'], element_xpath="//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    #政治面貌#(//input[@placeholder = '请选择字段项'])[6]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[7]")
    cursor.select_background(wait, driver, element_xpath="//div[@class = 'el-tree objectTree']", label='span', value=data_dict[gmsfzhm]['政治面貌'])
    #重新入党#(//input[@placeholder = '请选择'])
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择'])")
    cursor.select_background(wait, driver, element_xpath="//ul[@class = 'el-scrollbar__view el-select-dropdown__list']", label='span', value='否')
    #所属党组织#(//input[@placeholder = '请选择字段项'])[7]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[9]")
    cursor.select_party_org(wait, driver, tree=org_tree, target = data_dict[gmsfzhm]['接受申请党组织'], element_xpath="//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    #接受申请党组织#(//input[@placeholder = '请输入接受申请党组织'])[1]
    #cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入接受申请党组织'])[1]", text=data_dict[gmsfzhm]['接受申请党组织'])
    #一线情况#
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[8]")
    cursor.select_background(wait, driver, "//div[@class = 'el-tree objectTree']", label='span', value = data_dict[gmsfzhm]['一线情况'])
    #保存#(//span[contains(text(), '保存')])[1]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[1]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w=0.1, times=2)
    
    #编辑#(//span[contains(text(), '编辑')])[2]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[2]")
    #上传图片#//input[@type="file"]
    #需运行前验证文件夹是否存在
    current_working_dir = os.getcwd()
    absolute_path = current_working_dir + f"\\材料图片保存处\\{data_dict[gmsfzhm]['姓名']}+{gmsfzhm}\\入党申请书照片"
    cursor.upload_pic(pic_dir=absolute_path, input_xpath="//input[@type='file']", wait=wait, driver=driver)
    #保存#(//span[contains(text(), '保存')])[2]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[1]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w=0.1, times=2)
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(),'提交')])[2]", buttonxpath="(//span[contains(text(),'提交')])[2]", time_w=0.1, times=2)

    
    #警告#//div[@role = 'alert']
    #警告的内容#//div[@role = 'alert']//p[@class = 'el-message__content']



def step_2(driver, wait, data_dict, gmsfzhm):
    #编辑#(//span[contains(text(), '编辑')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[1]")
    #谈话人#(//input[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[1]", text=data_dict[gmsfzhm]['培养联系人'])
    #选择日期#(//input[@placeholder = '选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[1]", text=data_dict[gmsfzhm]['申请入党日期'])
    #保存#
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[1]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w=0.1, times=2)
    #党组织派人谈话记录#
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[1]")
    current_working_dir = os.getcwd()
    absolute_path = current_working_dir + f"\\材料图片保存处\\{data_dict[gmsfzhm]['姓名']}+{gmsfzhm}\\入党申请书照片"
    cursor.upload_pic(pic_dir=absolute_path, input_xpath="//input[@type='file']", wait=wait, driver=driver)
    #保存#
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[1]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w=0.1, times=2)
    #上报#//span[contains(text(),'上报')]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="//span[contains(text(),'上报')]", buttonxpath="//span[contains(text(),'上报')]", time_w=0.1, times=2)
    #再次进入该页面
    #cursor.commen_button(wait, driver, xpath="(//a[contains(text(), '查看详情')])[2]")
    #审核#(//span[contains(text(),'审核')])[2]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(),'审核')])[2]")
    #审核通过#(//span[contains(text(),'审核通过')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(),'审核通过')])[1]")


def step_3(driver, wait, data_dict, gmsfzhm, jg):
    #推荐和确定入党积极分子#(//a[contains(text(), '查看详情')])[3]
    #cursor.commen_button(wait, driver, xpath="(//a[contains(text(), '查看详情')])[3]")
    #1编辑#(//span[contains(text(), '编辑')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[1]")
    #党员推荐会议时间#(//input[@placeholder = '请选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请选择日期'])[1]", text=data_dict[gmsfzhm]['积极分子会议日期'])
    #上传双推会议纪录#(//input[@type="file"])[1]
    current_working_dir = os.getcwd()
    absolute_path = current_working_dir + f"\\材料图片保存处\\{data_dict[gmsfzhm]['姓名']}+{gmsfzhm}\\党员推荐会议记录"
    cursor.upload_pic(pic_dir=absolute_path, input_xpath="(//input[@type='file'])[1]", wait=wait, driver=driver)
    #保存#(//span[contains(text(), '保存')])[1]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[1]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w=0.1, times=2)

    #2编辑#(//span[contains(text(), '编辑')])[3]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[3]")
    #时间#(//input[@placeholder = '请选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请选择日期'])[1]",  text=data_dict[gmsfzhm]['积极分子会议日期'])
    t = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "(//input[@placeholder = '请选择日期'])[1]")
    t.send_keys(Keys.RETURN)
    #地点#(//input[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[1]", text='单位党务会议室')
    #应到人数#
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[3]", text=data_dict[gmsfzhm]["积极分子确定会议应到人数"])
    #实到人数#
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[4]", text=data_dict[gmsfzhm]["积极分子确定会议实到人数"])
    #缺席人数#
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[5]", text=data_dict[gmsfzhm]["积极分子确定会议应到人数"] - data_dict[gmsfzhm]["积极分子确定会议实到人数"])
    #会议类型#(//input[@placeholder = '请选择字段项'])[1]
    cursor.commen_button(wait, driver, xpath = "(//input[@placeholder = '请选择字段项'])[2]")
    cursor.select_background(wait, driver, "//div[@class = 'el-tree objectTree']", "span", "党员大会")
    #党支部书记#(//input[@placeholder = '请输入内容'])[3]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[2]", text=data_dict[gmsfzhm]['党支部书记'])
    #列席组织员#(//input[@placeholder = '请输入内容'])[4]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[6]", text=data_dict[gmsfzhm]['积极分子会议的列席组织员'])
    #会议研究意见#(//textarea[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//textarea[@placeholder = '请输入内容'])[1]", text='同意')
    #保存#(//span[contains(text(), '保存')])[3]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[3]", buttonxpath="(//span[contains(text(), '保存')])[3]", time_w=0.1, times=2)

    #3编辑#(//span[contains(text(), '编辑')])[4]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[4]")
    #研究确定入党积极分子会议纪录#(//input[@type="file"])[3]
    current_working_dir = os.getcwd()
    absolute_path = current_working_dir + f"\\材料图片保存处\\{data_dict[gmsfzhm]['姓名']}+{gmsfzhm}\\确定积极分子会议纪录"
    cursor.upload_pic(pic_dir=absolute_path, input_xpath="(//input[@type='file'])[3]", wait=wait, driver=driver)
    #保存#(//span[contains(text(), '保存')])[3]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[3]", buttonxpath="(//span[contains(text(), '保存')])[3]", time_w=0.1, times=2)
    
    #cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[5]")
    #籍贯#(//input[@placeholder = '请选择字段项'])[1]
    #cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[1]")

    #if jg.value == None:
    #    synchronizing_jg(driver=driver, wait=wait, input_node=jg)
    #jg.dayin()
    #cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[1]")
    #time.sleep(1)
    #cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[1]")
    #cursor.select_party_org(wait, driver, tree=jg, target=data_dict[gmsfzhm]['籍贯'], element_xpath = "//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    ##参加工作日期#(//input[@placeholder = '选择日期'])[2]
    #cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[2]", text=data_dict[gmsfzhm]['参加工作日期'])
    #
    #
    ##保存#(//span[contains(text(), '保存')])[4]
    #cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[4]", buttonxpath="(//span[contains(text(), '保存')])[4]", time_w=0.1, times=2)
    #上报#//span[contains(text(),'上报')]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="//span[contains(text(),'上报')]", buttonxpath="//span[contains(text(),'上报')]", time_w=0.1, times=2)
    #推荐和确定入党积极分子#(//a[contains(text(), '查看详情')])[3]
    #cursor.commen_button(wait, driver, xpath="(//a[contains(text(), '查看详情')])[3]")
    #审核#(//span[contains(text(),'审核')])[2]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(),'审核')])[2]")
    #审核通过#(//span[contains(text(),'审核通过')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(),'审核通过')])[1]")



def step_3_rebuild(driver, wait, data_dict, gmsfzhm, jg):
    #推荐和确定入党积极分子#(//a[contains(text(), '查看详情')])[3]
    #cursor.commen_button(wait, driver, xpath="(//a[contains(text(), '查看详情')])[3]")
    #编辑#(//span[contains(text(), '编辑')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[1]")
    #党员推荐会议时间#(//input[@placeholder = '请选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请选择日期'])[1]",  text=data_dict[gmsfzhm]['积极分子会议日期'])
    #上传双推会议纪录#(//input[@type="file"])[1]
    current_working_dir = os.getcwd()
    absolute_path = current_working_dir + f"\\材料图片保存处\\{data_dict[gmsfzhm]['姓名']}+{gmsfzhm}\\党员推荐会议记录"
    cursor.upload_pic(pic_dir=absolute_path, input_xpath="(//input[@type='file'])[1]", wait=wait, driver=driver)
    #保存#(//span[contains(text(), '保存')])[1]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[1]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w=0.1, times=2)

    #编辑#(//span[contains(text(), '编辑')])[3]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[3]")
    #时间#(//input[@placeholder = '请选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请选择日期'])[1]",  text=data_dict[gmsfzhm]['积极分子会议日期'])
    t = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "(//input[@placeholder = '请选择日期'])[1]")
    t.send_keys(Keys.RETURN)
    #地点#(//input[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[1]", text='单位党务会议室')
    #应到人数#
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[3]", text=data_dict[gmsfzhm]["积极分子确定会议应到人数"])
    #实到人数#
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[4]", text=data_dict[gmsfzhm]["积极分子确定会议实到人数"])
    #缺席人数#
    cursor.input_text(wait, driver, xpath="(//input[contains(@placeholder, '请输入内容')])[5]", text=data_dict[gmsfzhm]["积极分子确定会议应到人数"] - data_dict[gmsfzhm]["积极分子确定会议实到人数"])
    #会议类型#(//input[@placeholder = '请选择字段项'])[1]
    cursor.commen_button(wait, driver, xpath = "(//input[@placeholder = '请选择字段项'])[2]")
    cursor.select_background(wait, driver, "//div[@class = 'el-tree objectTree']", "span", "党员大会")
    
    
    #党支部书记#(//input[@placeholder = '请输入内容'])[3]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[2]", text=data_dict[gmsfzhm]['党支部书记'])
    #列席组织员#(//input[@placeholder = '请输入内容'])[4]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[6]", text=data_dict[gmsfzhm]['积极分子会议的列席组织员'])
    #会议研究意见#(//textarea[@placeholder = '请输入内容'])[1]
    cursor.input_text(wait, driver, xpath="(//textarea[@placeholder = '请输入内容'])[1]", text='同意')
    #保存#(//span[contains(text(), '保存')])[3]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[3]", buttonxpath="(//span[contains(text(), '保存')])[3]", time_w=0.1, times=2)

    #编辑#(//span[contains(text(), '编辑')])[4]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[4]")
    #研究确定入党积极分子会议纪录#(//input[@type="file"])[3]
    current_working_dir = os.getcwd()
    absolute_path = current_working_dir + f"\\材料图片保存处\\{data_dict[gmsfzhm]['姓名']}+{gmsfzhm}\\确定积极分子会议纪录"
    cursor.upload_pic(pic_dir=absolute_path, input_xpath="(//input[@type='file'])[3]", wait=wait, driver=driver)
    #保存#(//span[contains(text(), '保存')])[3]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[3]", buttonxpath="(//span[contains(text(), '保存')])[3]", time_w=0.1, times=2)
    #4编辑#(//span[contains(text(), '编辑')])[5]
    #工作单位及职务
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[5]")
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[1]", text=data_dict[gmsfzhm]['工作单位及职务'])


    #一线情况
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[4]")
    cursor.select_background(wait, driver, "//div[@class = 'el-tree objectTree']", label='span', value = data_dict[gmsfzhm]['一线情况'])

    #家庭住址
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入内容'])[2]", text=data_dict[gmsfzhm]['家庭住址'])

    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[4]", buttonxpath="(//span[contains(text(), '保存')])[4]", time_w=0.1, times=2)
    #cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[5]")
    #籍贯#(//input[@placeholder = '请选择字段项'])[1]
    #cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[1]")

    #if jg.value == None:
    #    synchronizing_jg(driver=driver, wait=wait, input_node=jg)
    #jg.dayin()
    #cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[1]")
    #time.sleep(1)
    #cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[1]")
    #cursor.select_party_org(wait, driver, tree=jg, target=data_dict[gmsfzhm]['籍贯'], element_xpath = "//div[@class = 'el-tree objectTree']", label_1 = 'span', label_2 = 'span')
    ##参加工作日期#(//input[@placeholder = '选择日期'])[2]
    #cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[2]", text=data_dict[gmsfzhm]['参加工作日期'])
    #
    #
    ##保存#(//span[contains(text(), '保存')])[4]
    #cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[4]", buttonxpath="(//span[contains(text(), '保存')])[4]", time_w=0.1, times=2)
    #上报#//span[contains(text(),'上报')]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="//span[contains(text(),'上报')]", buttonxpath="//span[contains(text(),'上报')]", time_w=0.1, times=2)
    #推荐和确定入党积极分子#(//a[contains(text(), '查看详情')])[3]
    #cursor.commen_button(wait, driver, xpath="(//a[contains(text(), '查看详情')])[3]")
    #审核#(//span[contains(text(),'审核')])[2]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(),'审核')])[2]")
    #审核通过#(//span[contains(text(),'审核通过')])[1]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(),'审核通过')])[1]")




def step_4(driver, wait, data_dict, gmsfzhm):
    #入党积极分子公示和备案#(//a[contains(text(), '查看详情')])[4]
    #cursor.commen_button(wait, driver, xpath="(//a[contains(text(), '查看详情')])[4]")
    #编辑#(//span[contains(text(), '编辑')])[2]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[2]")
    #公示起止日期#(//input[@placeholder = '开始日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '开始日期'])[1]", text=calculate_date_add(data_dict[gmsfzhm]['申请入党日期'],197))
    #公示起止日期#(//input[@placeholder = '结束日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '结束日期'])[1]", text=calculate_date_add(data_dict[gmsfzhm]['申请入党日期'],204))
    #公示结果#(//input[@placeholder = '请选择'])[1]
    cursor.commen_button(wait, driver, xpath="//span[@class = 'el-input__suffix-inner']")
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '通过')])[3]")

    #保存#(//span[contains(text(), '保存')])[2]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[2]", buttonxpath="(//span[contains(text(), '保存')])[2]", time_w=0.1, times=2)
    #编辑#(//span[contains(text(), '编辑')])[3]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '编辑')])[3]")
    #基层党委备案审查日期#(//input[@placeholder = '选择日期'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '选择日期'])[1]", text=calculate_date_add(data_dict[gmsfzhm]['申请入党日期'],197))
    t = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "(//input[@placeholder = '选择日期'])[1]")
    t.send_keys(Keys.RETURN) 
    #基层党委备案审查意见#(//input[@placeholder = '请选择字段项'])[1]
    cursor.commen_button(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[2]")
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(), '同意')])[1]")
    #意见#//textarea[@placeholder = '请输入内容']
    cursor.input_text(wait, driver, xpath="//textarea[@placeholder = '请输入内容']", text ="同意")
    #保存#(//span[contains(text(), '保存')])[2]
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(), '保存')])[2]", buttonxpath="(//span[contains(text(), '保存')])[2]", time_w=0.1, times=2)
    #提交#//span[contains(text(),'提交')]
    cursor.commen_button(wait, driver, xpath="(//span[contains(text(),'提交')])[2]")
    cursor.click_button_until_specifyxpath_disappear(wait=wait, driver=driver, specifyxpath="(//span[contains(text(),'提交')])[2]", buttonxpath="(//span[contains(text(),'提交')])[2]", time_w=0.1, times=2)
    pass


def step_5(driver, wait, data_dict, gmsfzhm):
    #指定培养联系人#(//a[contains(text(), '查看详情')])[5]
    #cursor.commen_button(wait, driver, xpath="(//a[contains(text(), '查看详情')])[5]")
    #添加#//span[contains(text(), '添加')]
    cursor.commen_button(wait, driver, xpath="//span[contains(text(), '添加')]")
    #现支部、其他支部#(//i[@class = 'el-select__caret el-input__icon el-icon-arrow-up'])[1]
    #选择联系人#(//i[@class = 'el-select__caret el-input__icon el-icon-arrow-up'])[1]#(//ul[@class = 'el-scrollbar__view el-select-dropdown__list'])[2]
    cursor.commen_button(wait, driver, xpath="(//i[@class = 'el-select__caret el-input__icon el-icon-arrow-up'])[2]")
    cursor.select_background_v1(wait, driver, element_xpath="(//ul[@class = 'el-scrollbar__view el-select-dropdown__list'])[2]", label='span', value=data_dict[gmsfzhm]["培养联系人"])
    #选择联系人（其他支部）#//input[@placeholder='请输入']
    #入党日期(现支部不需)#(//input[@placeholder='选择日期'])[1]
    #学历(现支部不需)#(//input[@placeholder='请选择字段项'])
    #自何时起负责培养#(//input[@placeholder='选择日期'])[2]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder='选择日期'])[2]", text=calculate_date_add(data_dict[gmsfzhm]['申请入党日期'],198))
    cursor.input_text(wait, driver, xpath="(//input[@placeholder='选择日期'])[2]", text="\n")
    #保存#(//span[contains(text(), "保存")])[1]
    cursor.click_button_until_specifyxpath_disappear(wait, driver, specifyxpath="(//span[contains(text(), '保存')])[1]", buttonxpath="(//span[contains(text(), '保存')])[1]", time_w= 0.8, times=2)
    #提交#(//span[contains(text(), "提交")])[1]
    cursor.click_button_until_specifyxpath_disappear(wait, driver, specifyxpath="(//span[contains(text(), '提交')])[1]", buttonxpath="(//span[contains(text(), '提交')])[1]", time_w= 0.8, times=2)







if __name__ == "__main__":
    main()