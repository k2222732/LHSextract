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

def jibenxinxi_download(file, path, wait, countx, count, wb, ws):
    #填写序号
    yi = count
    ws.cell(row = countx, column = 1, value =yi)
    #姓名
    name_temp = wait.until(EC.visibility_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
    er = name_temp
    ws.cell(row = countx, column = 2, value =er)
    #循环断言
    while(1):
        try:
            df = ws.cell(row=countx, column=2).value
            assert df != ""
            break
        except AssertionError:
            time.sleep(0.5)
            er = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[2]/span"))).text
            ws.cell(row = countx, column = 2, value =er)
    #性别
    san = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[2]//span[1]"))).text
    ws.cell(row = countx, column = 3, value =san)
    #公民身份证号码
    si = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[4]//span[1]"))).text
    ws.cell(row = countx, column = 4, value =si)
    #民族
    wu = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[1]/td[6]//span[1]"))).text
    ws.cell(row = countx, column = 5, value =wu)
    #出生日期
    liu = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[2]/td[4]//span[1]"))).text
    ws.cell(row = countx, column = 6, value =liu)
    #学历
    bitian_located(file, path, countx, column=7, wait = wait, str = "(//tbody)[1]/tr[2]/td[6]//span[1]", wb=wb, ws=ws)
    #申请入党日期
    ba = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[2]//span[1]"))).text
    ws.cell(row = countx, column = 8, value =ba)
    #手机号码
    jiu = wait.until(EC.presence_of_element_located((By.XPATH, "(//tbody)[1]/tr[3]/td[4]//span[1]"))).text
    ws.cell(row = countx, column = 9, value =jiu)
    #背景信息
    bitian_located(file, path, countx, column=10, wait = wait, str = "(//tbody)[1]/tr[3]/td[6]//span[1]", wb=wb, ws=ws)
    #工作岗位
    bitian_located(file, path, countx, column=11, wait = wait, str = "(//tbody)[1]/tr[4]/td[2]//span[1]", wb=wb, ws=ws)
    #政治面貌
    bitian_located(file, path, countx, column=12, wait = wait, str = "(//tbody)[1]/tr[4]/td[4]//span[1]", wb=wb, ws=ws)
    #接受申请党组织
    shisan = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '接收入党申请的党组织')])[1]//following-sibling::td").get_attribute("textContent")
    ws.cell(row = countx, column = 13, value =shisan)
    wb.save(path)
    return name_temp


def jijifenzi_download(path, file, wait, countx, control, wb, ws):
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
        shiwu = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[4]//span)[1]"))).text
        ws.cell(row = countx, column = 15, value =shiwu)
        #参加工作日期
        shiliu = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[4]/td[6]//span)[1]"))).text
        ws.cell(row = countx, column = 16, value =shiliu)
        #申请入党日期
        shiqi = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[2]//span)[1]"))).text
        ws.cell(row = countx, column = 17, value =shiqi)
        #确定入党积极分子日期
        shiba = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[5]/td[4]//span)[1]"))).text
        ws.cell(row = countx, column = 18, value =shiba)
        #工作单位及职务
        shijiu = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[6]/td[4]//span)[1]"))).text
        ws.cell(row = countx, column = 19, value =shijiu)
        #家庭住址
        ershi = wait.until(EC.presence_of_element_located((By.XPATH, "((//tbody)[2]/tr[7]/td[2]//span)[1]"))).text
        ws.cell(row = countx, column = 20, value =ershi)
        
        if control == 3 or 4:
            ershiyi = "-"
            ws.cell(row = countx, column = 21, value =ershiyi)

            ershisan = "-"
            ws.cell(row = countx, column = 23, value =ershisan)

            ershisi = "-"
            ws.cell(row = countx, column = 24, value =ershisi)

            ershiwu = "-"
            ws.cell(row = countx, column = 25, value =ershiwu)

            ws.cell(row=countx, column=26).value = ws.cell(row=countx, column=28).value
        
        elif control == 1:
            pass
        elif control == 5:
            pass

        # 人员类型
        if control == 3:
            ershiqi = "发展对象"
            ws.cell(row = countx, column = 27, value =ershiqi)
            
        elif control == 4:
            ershiqi = "积极分子"
            ws.cell(row = countx, column = 27, value =ershiqi)
        
        wb.save(path)


def yubeidangyuan(path, file, wait, countx, control, wb, ws):
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
                
        ershiyi = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '加入党组织日期')])[1]//following-sibling::td[1]").get_attribute("textContent")
        ws.cell(row = countx, column = 21, value =ershiyi)
        # 转为正式党员日期
        ershier = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '转为正式党员日期')])[1]//following-sibling::td[1]").get_attribute("textContent")
        ws.cell(row = countx, column = 22, value =ershier)
        # 人员类别
        ershisan = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '人员类别')])[1]//following-sibling::td[1]").get_attribute("textContent")
        ws.cell(row = countx, column = 23, value =ershisan)
        # 党籍状态
        ershisi = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '党籍状态')])[1]//following-sibling::td[1]").get_attribute("textContent")
        ws.cell(row = countx, column = 24, value =ershisi)
        # 入党类型
        ershiwu = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '入党类型')])[1]//following-sibling::td[1]").get_attribute("textContent")
        ws.cell(row = countx, column = 25, value =ershiwu)
        # 所在党支部
        ershiliu = wait_return_subelement_absolute(wait, time_w = 0.5, xpath="(//td[contains(text(), '所在党支部')])[1]//following-sibling::td[1]").get_attribute("textContent")
        ws.cell(row = countx, column = 26, value =ershiliu)
        # 人员类型
        if control == 1:
            ershiqi = "正式党员"
            ws.cell(row = countx, column = 27, value =ershiqi)
        elif control == 2:
            ershiqi = "预备党员"
            ws.cell(row = countx, column = 27, value =ershiqi)
        wb.save(path)
        

def bitian_located(file, path, row, column, wait, str, wb, ws):
    z = wait_return_subelement_absolute(wait, time_w=0.5, xpath=str).text
    ws.cell(row=row, column=column, value=z)
    #file.active.cell(row=row, column=column).value = wait.until(EC.presence_of_element_located((By.XPATH, str))).text
    start_time = time.time()
    while(1):
            elapsed_time = time.time() - start_time
            if elapsed_time > 5:
                break
            try:
                df = ws.cell(row=row, column=column).value
                assert df != ""
                break
            except AssertionError:
                z = wait.until(EC.presence_of_element_located((By.XPATH, str))).text
                ws.cell(row=row, column=column, value=z)
                