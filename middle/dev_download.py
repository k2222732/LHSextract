from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from datetime import datetime
import pandas as pd
import time 
import os
import re
import configparser
import threading
from bs4 import BeautifulSoup
from base.waitclick import *
from base.write_entry import *

def jibenxinxi_download(file, path, wait, countx, count):
    
    #填写序号
    file.active.cell(row=countx, column=1).value = count
    file.save(path)
    
    #姓名
    name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
    file.active.cell(row=countx, column=2).value = name_temp
    file.save(path)
    #循环断言
    while(1):
        try:
            df = file.active.cell(row=countx, column=2).value
            assert df != ""
            break
        except AssertionError:
            time.sleep(0.5)
            file.active.cell(row=countx, column=2).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
            file.save(path)
    #性别
    file.active.cell(row=countx, column=3).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[2]//span[1]"))).text
    
    #公民身份证号码
    file.active.cell(row=countx, column=4).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[4]//span[1]"))).text
    
    #民族
    file.active.cell(row=countx, column=5).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[6]//span[1]"))).text

    #出生日期
    file.active.cell(row=countx, column=6).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[4]//span[1]"))).text

    #学历
    bitian_located(file, path, countx, column=7, wait = wait, str = "(//tbody)[1]/tr[2]/td[6]//span[1]")
    #申请入党日期
    file.active.cell(row=countx, column=8).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[2]//span[1]"))).text
    
    #手机号码
    file.active.cell(row=countx, column=9).value = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[4]//span[1]"))).text
    #背景信息
    bitian_located(file, path, countx, column=10, wait = wait, str = "(//tbody)[1]/tr[3]/td[6]//span[1]")
    #工作岗位
    bitian_located(file, path, countx, column=11, wait = wait, str = "(//tbody)[1]/tr[4]/td[2]//span[1]")
    #政治面貌
    bitian_located(file, path, countx, column=12, wait = wait, str = "(//tbody)[1]/tr[4]/td[4]//span[1]")
    #接受申请党组织
    file.active.cell(row=countx, column=13).value = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '接收入党申请的党组织')])[1]//following-sibling::td").get_attribute("textContent")
    file.save(path)
    return name_temp


def jijifenzi_download(path, file, wait, countx, control):
    #切换到积极分子信息
        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '入党积极分子基本信息')]")))
        # #籍贯
        # while 1:
        #     jiguan = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[2]/tr[3]/td[4]//span[1]"))).text
        #     if jiguan == "":
        #         pass
        #     else:
        #         file.active.cell(row=countx, column=14).value = jiguan
        #         break
        #入团日期
        file.active.cell(row=countx, column=15).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[4]//span)[1]"))).text
        
        #参加工作日期
        file.active.cell(row=countx, column=16).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[6]//span)[1]"))).text
        
        #申请入党日期
        file.active.cell(row=countx, column=17).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[2]//span)[1]"))).text
        
        #确定入党积极分子日期
        file.active.cell(row=countx, column=18).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[4]//span)[1]"))).text
        
        #工作单位及职务
        file.active.cell(row=countx, column=19).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[6]/td[4]//span)[1]"))).text
        
        #家庭住址
        file.active.cell(row=countx, column=20).value = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[7]/td[2]//span)[1]"))).text
        file.save(path)
        if control == 3 or 4:
            file.active.cell(row=countx, column=21).value = "-"
            file.active.cell(row=countx, column=23).value = "-"
            file.active.cell(row=countx, column=24).value = "-"
            file.active.cell(row=countx, column=25).value = "-"
            file.active.cell(row=countx, column=26).value = file.active.cell(row=countx, column=28).value
        elif control == 1:
            pass
        elif control == 5:
            pass

        # 人员类型
        if control == 3:
            file.active.cell(row=countx, column=27).value = "发展对象"
            file.save(path)
        elif control == 4:
            file.active.cell(row=countx, column=27).value = "积极分子"
            file.save(path)
        


def yubeidangyuan(path, file, wait, countx, control):
    #切换到预备党员
        while(1):
            try:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '预备党员')]")))
                break
            except:
                time.sleep(0.1)
        while(1):
            try:
                company_info.click()
                break
            except:
                company_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '预备党员')]")))
        # 加入党组织日期
                
        file.active.cell(row=countx, column=21).value = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '加入党组织日期')])[1]//following-sibling::td[1]").get_attribute("textContent")
        
        # 转为正式党员日期
        file.active.cell(row=countx, column=22).value = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '转为正式党员日期')])[1]//following-sibling::td[1]").get_attribute("textContent")
               
        # 人员类别
        file.active.cell(row=countx, column=23).value = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '人员类别')])[1]//following-sibling::td[1]").get_attribute("textContent")
               
        # 党籍状态
        file.active.cell(row=countx, column=24).value = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '党籍状态')])[1]//following-sibling::td[1]").get_attribute("textContent")
           
        # 入党类型
        file.active.cell(row=countx, column=25).value = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '入党类型')])[1]//following-sibling::td[1]").get_attribute("textContent")
        
        # 所在党支部
        file.active.cell(row=countx, column=26).value = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '所在党支部')])[1]//following-sibling::td[1]").get_attribute("textContent")
        
        # 人员类型
        if control == 1:
            file.active.cell(row=countx, column=27).value = "正式党员"
            file.save(path)
        elif control == 2:
            file.active.cell(row=countx, column=27).value = "预备党员"
            file.save(path)
        

def bitian_located(file, path, row, column, wait, str):
    file.active.cell(row=row, column=column).value = wait_return_subelement_absolute(wait, time_w=0.5, xpath=str).text
    #file.active.cell(row=row, column=column).value = wait.until(EC.presence_of_element_located((By.XPATH, str))).text
    start_time = time.time()
    while(1):
            elapsed_time = time.time() - start_time
            if elapsed_time > 5:
                break
            try:
                df = file.active.cell(row=row, column=column).value
                assert df != ""
                break
            except AssertionError:
                file.active.cell(row=row, column=column).value = wait.until(EC.presence_of_element_located((By.XPATH, str))).text
                file.save(path)