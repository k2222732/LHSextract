#上传灯塔大课堂,党委批量上传
#上传灯塔大课堂，支部自己上传

#检查谁没上传主题党日、灯塔大课堂（仅党委有此功能）
import dev_func
from base.waitclick import wait_click_xpath, wait_return_element_text, wait_return_subelement_absolute_notmust,wait_return_subelement_absolute
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
from base.solv_date import calculate_date_add
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import dev_func
from bs4 import BeautifulSoup
import middle.switch_role
import time
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from selenium.webdriver.support import expected_conditions as EC
from base.mystruct import TreeNode
from base.waitclick import wait_click_xpath, wait_click_xpath_relative, wait_return_subelement_relative, wait_return_subelement_absolute

hdzt = ""
hdsj = ""
text = ""
zbordw = ""
qc = ""



def main():
    driver = login._main()
    wait = WebDriverWait(driver, 2, 0.5)
    dev_func.access_dev_database(driver, wait)
    dev_func.switch_role(wait, driver)
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
    middle.switch_role.switch_role(wait)
    #//span[contains(text(), '党务管理')]
    #(//span[contains(text(), '支部活动')])[1]
    #(//span[contains(text(), '记入电子日志')])[1]
    #wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[8]")
    #//input[@placeholder = '请输入活动主题'] <--- hdzt   
    #//input[@class = 'vue-treeselect__input']    出现
    #//div[@class = 'vue-treeselect__option-arrow-container']   展开
    #(//div[@class = 'vue-treeselect__list'])[2]     获取
    #select_party_org(wait, driver, tree:TreeNode, target:str, element_xpath, label_1 = 'span', label_2 = 'span'):
    #//span[contains(text(), '支部党员大会')]
    #//span[contains(text(), '支部委员会')]
    #//span[contains(text(), '支部委员会') and @class = 'fs-checkbox__label']
    #//span[contains(text(), '党课') and @class = 'fs-checkbox__label']
    #//span[contains(text(), '灯塔大课堂')] 
    #//input[@placeholder = '请选择期次']   (//ul[@class = 'fs-scrollbar__view fs-select-dropdown__list'])[1]    (//ul[@class = 'fs-scrollbar__view fs-select-dropdown__list'])[1]//span[contains(text(), qc)]  
    #//input[@placeholder = '请输入主持人']  <--  (//div[@aria-label = 'checkbox-group'])[2]/label[1]
    #//input[@placeholder = '请输入记录人']  <--  (//div[@aria-label = 'checkbox-group'])[2]/label[2]
    #(//label[@class = 'fs-checkbox left-checkbox margin-right-10-px']//span[@class = 'fs-checkbox__input'])[1]   全选
    #driver.switch_to.frame(driver.find_element(By.ID, "edui1_iframeholder"))
    #//html[@class = 'view']//body  <--  text
    #driver.switch_to.default_content()
    #//span[contains(text(), '保存并归档')]
    #//button[@class = 'el-button el-button--default el-button--small el-button--primary ']
    input("")


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
        if div_element and div_element.get('class') == 'is-leaf el-tree-node__expand-icon el-icon-caret-right':
            next
        else:
            wait_click_xpath_relative(wait=wait,time_w = 0.5, element = item, xpath = ".el-tree-node__expand-icon el-icon-caret-right")
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
