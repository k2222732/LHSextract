import time
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from selenium.webdriver.support import expected_conditions as EC
from base.mystruct import TreeNode
from base.waitclick import *



def synchronizing_org(wait, driver, input_node:TreeNode):
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[8]")
    root_org = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "(//div[@class = 'el-tree objectTree']/div[@role = 'treeitem']/div[@class = 'el-tree-node__content'])")
    #采集根组织信息
    root_text = root_org.get_attribute('textContent')
    input_node.set_value(value=root_text)
    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree objectTree']/div[@role = 'treeitem']/div[@role = 'group']")
    recursion_org(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[8]")

def synchronizing_org_v1(wait, driver, input_node:TreeNode, xpath, xpath2, xpath3, xpath4):
    
    index = 1
    root_org = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = xpath)
    #采集根组织信息
    root_text = root_org.get_attribute('textContent')
    input_node.set_value(value=root_text)
    input_node.index = index
    index = index +1
    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = xpath2)
    time.sleep(3)
    recursion_org_v1(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node, index=index, xpath3=xpath3, xpath4=xpath4)



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
        if div_element and 'is-leaf' in div_element.get('class'):
            next
        else:
            wait_click_xpath_relative(wait=wait,time_w = 0.5, element = item, xpath = ".//div[@class = 'el-tree-node__content']/span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            new_tree_root = wait_return_subelement_relative(time_w = 0.5, element= item, xpath = ".//div[@role = 'group']")
            recursion_org(new_tree_root, wait, driver, new_node)


def recursion_org_v1(tree_root, wait, driver, root_node:TreeNode, index, xpath3, xpath4):
    
    org_items = []

    org_items = wait_return_subelement_relative_v1(1, tree_root, xpath=".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        new_node.index = index
        index = index +1
        root_node.add_child(child_node=new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('span')
        if div_element and 'is-leaf' in div_element.get('class'):
            next
        else:
            wait_click_xpath_relative(wait=wait,time_w = 0.5, element = item, xpath = xpath3)
            new_tree_root = wait_return_subelement_relative(time_w = 0.5, element= item, xpath = xpath4)
            index = recursion_org_v1(new_tree_root, wait, driver, new_node, index, xpath3=xpath3, xpath4=xpath4)
    return index

def synchronizing_job(wait, driver, input_node:TreeNode):
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[6]")
    #采集根组织信息
    root_text = 'root'
    input_node.set_value(value=root_text)
    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree objectTree']")
    recursion_job(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[6]")



def recursion_job(tree_root, wait, driver, root_node:TreeNode):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        root_node.add_child(child_node=new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('div')
        if div_element and div_element.get('aria-disabled') == 'true':
            wait_click_xpath_relative(wait=wait,time_w = 0.5, element = item, xpath = ".//div[@class = 'el-tree-node__content']/span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            time.sleep(0.5)
            new_tree_root = wait_return_subelement_relative(time_w = 0.5, element= item, xpath = ".//div[@role = 'group']")
            recursion_job(new_tree_root, wait, driver, new_node)
        else:
            next





def synchronizing_xueli(wait, driver, input_node:TreeNode):
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[3]")
    #采集根组织信息
    root_text = 'root'
    input_node.set_value(value=root_text)
    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree objectTree']")
    recursion_xueli(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[3]")



def recursion_xueli(tree_root, wait, driver, root_node:TreeNode):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        root_node.add_child(child_node=new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('div')
        if div_element and div_element.get('aria-disabled') == 'true':
            wait_click_xpath_relative(wait=wait,time_w = 0.5, element = item, xpath = ".//div[@class = 'el-tree-node__content']/span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            time.sleep(0.3)
            new_tree_root = wait_return_subelement_relative(time_w = 0.5, element= item, xpath = ".//div[@role = 'group']")
            recursion_xueli(new_tree_root, wait, driver, new_node)
        else:
            next



def synchronizing_jg(wait, driver, input_node:TreeNode):
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[3]")
    #采集根组织信息
    root_text = 'root'
    input_node.set_value(value=root_text)
    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree objectTree']")
    recursion_jg(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)
    wait_click_xpath(wait, driver, xpath="(//input[@placeholder = '请选择字段项'])[3]")



def recursion_jg(tree_root, wait, driver, root_node:TreeNode):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        root_node.add_child(child_node=new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('div')
        if div_element and div_element.get('aria-disabled') == 'true':
            wait_click_xpath_relative(wait=wait,time_w = 0.5, element = item, xpath = ".//div[@class = 'el-tree-node__content']/span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            time.sleep(0.3)
            new_tree_root = wait_return_subelement_relative(time_w = 0.5, element= item, xpath = ".//div[@role = 'group']")
            recursion_jg(new_tree_root, wait, driver, new_node)
        else:
            next