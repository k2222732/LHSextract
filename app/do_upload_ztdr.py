#上传灯塔大课堂,党委批量上传
#上传灯塔大课堂，支部自己上传

#检查谁没上传主题党日、灯塔大课堂（仅党委有此功能）
import dev_func
from base.waitclick import *
from selenium.webdriver.support.ui import WebDriverWait
import login
import os
import pandas as pd
from tkinter import messagebox
import base.rwconfig as conf
import base.write_entry as cursor
from base.boot import *
import time
from base.solv_date import calculate_date_add
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import dev_func
from bs4 import BeautifulSoup
import middle.switch_role
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from selenium.webdriver.support import expected_conditions as EC
from base.mystruct import TreeNode
from base.waitclick import *
from base.write_entry import *
from selenium.common.exceptions import NoSuchElementException
import json
from typing import List


hdzt = "四月份主题党日"#活动主题
hdsj = {"开始时间":"2024-04-29 00:00", "结束时间":"2024-04-29 00:30"}#活动时间
text = ("开展“不忘初心，牢记使命”主题党日活动\n"
        "按照上级党组织要求,我支部开展“不忘初心，牢记使命”主题党日活动。支部全体党员参加会议，会议由党支部书记主持。\n"
       " 1. 坚持第一议题。学习习近平总书记关于本领域、本行业、本部门的最新讲话精神，认真学习习近平总书记在青海、宁夏考察、中央军委政治工作会议上的重要讲话和对防汛抗旱工作作出的重要指示精神，掌握其中蕴含的领导方法、思想方法、工作方法，不断提高履职能力和水平。\n"
        "2. 持续抓好党纪学习教育。\n"
        "3. 做好党员发展工作。\n"
        "4. 填报6月份《积分手册》经党组织审核后公示并上报每名党员积分情况。\n"
        "5. 收缴本月党费，做好上半年发展党员、入党积极分子培养工作。\n"
        "通过学习坚定了全体党员的理想信念，增强了党性意识，提高了责任担当，为充分发挥党员先锋模范作用、砥砺前行奠定了扎实基础。")
#三会一课内容
zbordw = ""#支部还是党委
qc = "202406  “请党放心 强国有我”——青年党员学习研讨"#期次
excel_path = "G:\project\LHSextract\三会一课未上传\三会一课未上传名单.xlsx"#未上传名单的路径


def init(args:list[str]):
    global hdzt
    global hdsj
    global text
    hdzt = args[0]
    hdsj = args[1]
    json_str = "{" + hdsj + "}"
    # 将 JSON 字符串转换为字典
    try:
        hdsj = json.loads(json_str)
        print(hdsj)
        # 输出: {'开始时间': '2024-07-01 00:00', '结束时间': '2024-07-01 00:30'}
    except json.JSONDecodeError as e:
        print(f"JSONDecodeError: {e}")
    text = args[2]



def main(args:list[str]):
    init(args)
    result = []
    driver = login._main()
    wait = WebDriverWait(driver, 2, 0.5)
    dev_func.access_dev_database(driver, wait)
    dev_func.switch_role(wait, driver)
    commen_button(wait, driver, xpath="(//span[contains(text(), '人员信息')])[1]")
    #(//span[contains(text(), '人员信息')])[1]

    org_tree = TreeNode()
    synchronizing_org(wait, driver, input_node=org_tree)
    try:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if '管理员工作台' in driver.title:
                break
        print(f"返回管理员工作台成功")
    except:
        time.sleep(1)
    middle.switch_role.access_e_shandong(driver, wait)
    middle.switch_role.switch_role_v1(wait)
    #//span[contains(text(), '党务管理')]
    commen_button(wait, driver, xpath="//span[contains(text(), '党务管理')]")
    #(//span[contains(text(), '支部活动')])[1]
    commen_button(wait, driver, xpath="(//span[contains(text(), '支部活动')])[1]")
    df = pd.read_excel(excel_path)
    data_dict = {}
    for index, row in df.iterrows():#df.iterrows()方法返回一个迭代器,该迭代器会生成DataFrame中每行的索引和行数据。具体来说,每次迭代时,它会返回一个包含两个元素的元组
            dkt = row['未上传主题党日的企业']
            nested_dict = {'未上传主题党日的企业':dkt}
            data_dict[index]=nested_dict
    for qy in data_dict:
        target = data_dict[qy]['未上传主题党日的企业']
        ####################循环开始#########################
        #(//span[contains(text(), '记入电子日志')])[1]
        commen_button(wait, driver, xpath="(//span[contains(text(), '记入电子日志')])[1]")
        #wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[8]")
        #//input[@placeholder = '请输入活动主题'] <--- hdzt 
        input_text(wait, driver, xpath="//input[@placeholder = '请输入活动主题']", text=hdzt)
        time.sleep(1)
        input_text(wait, driver, xpath="//input[@placeholder = '请输入活动主题']", text=hdzt)
        #//input[@class = 'vue-treeselect__input']    出现
        commen_button(wait, driver, xpath="//input[@class = 'vue-treeselect__input']")
        #//div[@class = 'vue-treeselect__option-arrow-container']   展开
        commen_button(wait, driver, xpath="//div[@class = 'vue-treeselect__option-arrow-container']")
        #(//div[@class = 'vue-treeselect__list'])[2]     获取
        element_list = wait_return_subelement_absolute(wait, time_w=0.5, xpath="(//div[@class = 'vue-treeselect__list'])[2]")
        #select_party_org(wait, driver, tree:TreeNode, target:str, element_xpath, label_1 = 'span', label_2 = 'span'):
        select_party_org(wait, driver, tree=org_tree, target=target, element_xpath="(//div[@class = 'vue-treeselect__list'])[2]", label_1 = 'label', label_2 = 'label')
        #//span[contains(text(), '支部党员大会')]
        #commen_button(wait, driver, xpath="//span[contains(text(), '支部党员大会')]")
        #//span[contains(text(), '支部委员会')]
        #commen_button(wait, driver, xpath= "//span[contains(text(), '支部委员会')]")
        #//span[contains(text(), '支部委员会') and @class = 'fs-checkbox__label']
        commen_button(wait, driver, xpath="//span[contains(text(), '支部委员会') and @class = 'fs-checkbox__label']")
        commen_button(wait, driver, xpath="//span[contains(text(), '主题党日') and @class = 'fs-checkbox__label']")
        #//span[contains(text(), '党课') and @class = 'fs-checkbox__label']
        #commen_button(wait, driver, xpath="//span[contains(text(), '党课') and @class = 'fs-checkbox__label']")
        #//span[contains(text(), '灯塔大课堂')] 
        #commen_button(wait, driver, xpath="//span[contains(text(), '灯塔大课堂')]")
        #//input[@placeholder = '请选择期次']   (//ul[@class = 'fs-scrollbar__view fs-select-dropdown__list'])[1]    (//ul[@class = 'fs-scrollbar__view fs-select-dropdown__list'])[1]//span[contains(text(), qc)]  
        #commen_button(wait, driver, xpath="//input[@placeholder = '请选择期次']")
        #time.sleep(1)
        #try:
        #    temp = driver.find_element(By.XPATH, f"(//ul[@class = 'fs-scrollbar__view fs-select-dropdown__list'])[1]//span[contains(text(), '{qc}')]")
        #except NoSuchElementException as e:
        #    result.append(f"{data_dict[qy]['未上传大课堂的企业']}已上传\n")
        #    driver.back()
        #    continue
        #select_background(wait, driver, element_xpath="(//ul[@class = 'fs-scrollbar__view fs-select-dropdown__list'])[1]", label='span', value=qc)


        #//input[@placeholder = '请输入主持人']  <--  (//div[@aria-label = 'checkbox-group'])[2]/label[1]
        element_zcr = wait.until(EC.presence_of_element_located((By.XPATH, "(//div[@aria-label = 'checkbox-group'])[2]/label[1]")))
        zcr = element_zcr.get_attribute('textContent')
        input_text(wait, driver, xpath="//input[@placeholder = '请输入主持人']", text=zcr)
        #//input[@placeholder = '请输入记录人']  <--  (//div[@aria-label = 'checkbox-group'])[2]/label[2]
        element_jlr = wait.until(EC.presence_of_element_located((By.XPATH, "(//div[@aria-label = 'checkbox-group'])[2]/label[2]")))
        jlr = element_jlr.get_attribute('textContent')
        input_text(wait, driver, xpath="//input[@placeholder = '请输入记录人']", text=jlr)
        #//input[@placeholder = '开始时间']
        input_text_v1(wait, driver, xpath="//input[@placeholder = '开始时间']", text=hdsj['开始时间']+'\n', times=3)
        input_text_v1(wait, driver, xpath="//input[@placeholder = '结束时间']", text=hdsj['结束时间']+'\n', times=3)
        commen_button(wait, driver, xpath="(//span[contains(text(), '确定')]/..)[2]")
        #//input[@placeholder = '结束时间']
        
        
        #//input[@placeholder = '请输入活动地点']
        input_text(wait, driver, xpath="//input[@placeholder = '请输入活动地点']", text="单位党务活动室")
        #(//label[@class = 'fs-checkbox left-checkbox margin-right-10-px']//span[@class = 'fs-checkbox__input'])[1]   全选
        commen_button(wait, driver, xpath="(//label[@class = 'fs-checkbox left-checkbox margin-right-10-px']//span[@class = 'fs-checkbox__input'])[1]")
        #driver.switch_to.frame(driver.find_element(By.ID, "edui1_iframeholder"))
        while(1):
           try:
               iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//iframe[1]")))
               driver.switch_to.frame(iframe)
               break
           except:
               time.sleep(1)
        # //html[@class = 'view']//body  <--  text
        word = wait_return_subelement_absolute(wait, time_w = 2, xpath="//html[@class = 'view']//body")
        word.send_keys(text)
        #driver.switch_to.default_content()
        driver.switch_to.default_content()
        #//span[contains(text(), '保存并归档')]
        #button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '保存并归档')]/..")))
        #button.click()
        time.sleep(1)
        commen_button(wait, driver, xpath="//span[contains(text(), '保存并归档')]/..")
        #//button[@class = 'el-button el-button--default el-button--small el-button--primary ']
        #commen_button(wait, driver, xpath="//button[@class = 'el-button el-button--default el-button--small el-button--primary ']")
        time.sleep(1)
        click_button_until_specifyxpath_disappear(wait, driver, specifyxpath="//button[@class = 'el-button el-button--default el-button--small el-button--primary ']", buttonxpath="//button[@class = 'el-button el-button--default el-button--small el-button--primary ']", time_w = 1, times = 3)
        ####################循环结束#########################
    for element in result:
        print(element)
    input("按任意键结束")


def synchronizing_org(wait, driver, input_node:TreeNode):
    #wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[8]")
    root_org = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "(//span[@class = 'el-tree-node__label'])[1]")
    #采集根组织信息
    root_text = root_org.get_attribute('textContent')
    input_node.set_value(value=root_text)
    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree-node__children']")
    recursion_org(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)


    
    


    
    

def recursion_org(tree_root, wait, driver, root_node:TreeNode):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        root_node.add_child(child_node=new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('span')
        test = div_element.get('class')
        if div_element and test[0] == 'is-leaf':
            next
        else:
            wait_click_xpath_relative(wait=wait,time_w = 0.5, element = item, xpath = ".//span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            new_tree_root = wait_return_subelement_relative(time_w = 0.5, element= item, xpath = ".//div[@role = 'group']")
            recursion_org(new_tree_root, wait, driver, new_node)
            



#点选列表（树）                                                                          label             label
def select_party_org(wait, driver, tree:TreeNode, target:str, element_xpath, label_1 = 'span', label_2 = 'span'):
    '''
    target:最终目标的叶子节点，通过该节点回溯到根节点然后依次点击到达target节点
    element_xpath:列表的xpath
    '''
    target_node = tree.find_node(target)
    path_to_root = target_node.path_to_root()
    path_to_root = path_to_root[1:]
    current_element = wait_return_subelement_absolute(wait, time_w=0.5, xpath=element_xpath)
    for i, node_value in enumerate(path_to_root): #enumerate(path_to_root) 会产生 [(0, 'node1'), (1, 'node2'), (2, 'node3')]
        #如果node_value不是列表中的最后一个元素，则点击下拉箭头
                                    #(//div[@role = 'treeitem'])[1]//span[contains(text(),'中国共产党山东汶上经济开发区工作委员会')]
        if i<len(path_to_root)-1:
            wait_click_xpath_relative(wait, time_w = 0.5, element = current_element, xpath = f".//{label_1}[contains(text(), '{node_value}')]/../preceding-sibling::div")
                                                                    #(//div[@role = 'treeitem'])[1]//span[contains(text(), '中国共产党山东汶上经济开发区工作委员会')]/../../following-sibling::div[@role = 'group']
            current_element = wait_return_subelement_relative(time_w = 0.5, element = current_element, xpath = f".//{label_1}[contains(text(), '{node_value}')]/../../following-sibling::div[@class = 'vue-treeselect__list']")
        #如果node_value是列表中的最后一个元素，则点击元素体
        else:                                                #测试：  (//div[@role = 'treeitem'])[1]//span[contains(text(), '中国共产党山东汶上经济开发区工作委员会')]/../../following-sibling::div[@role = 'group']/.//span[contains(text(), '中国共产党新风光电子科技股份有限公司委员会')]/../preceding-sibling::span
            wait_click_xpath_relative(wait, time_w = 0.5, element = current_element, xpath = f".//{label_2}[contains(text(), '{node_value}')]")                                                                                                     #测试： (//div[@role = 'treeitem'])[1]//span[contains(text(), '中国共产党山东汶上经济开发区工作委员会')]/../../following-sibling::div[@role = 'group']//span[contains(text(), '中国共产党新风光电子科技股份有限公司第一支部委员会')]


if __name__ == "__main__":
    main()
