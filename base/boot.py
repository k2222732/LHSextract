import time
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from selenium.webdriver.support import expected_conditions as EC
from mystruct import TreeNode
from waitclick import wait_click_xpath, wait_click_xpath_relative, wait_return_subelement_relative, wait_return_subelement_absolute



def synchronizing_org(wait, driver, input_node):
    root_org = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "(//div[@class = 'el-tree objectTree']/div[@role = 'treeitem']/div[@class = 'el-tree-node__content'])")
    #采集根组织信息
    root_text = root_org.get_attribute('textContent')
    input_node.set_value(root_text)
    #获取根节点结构体//div[@class = "tree_wrapper"]//div[@role = "treeitem"]/div[@role = "group"]到tree_root
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree objectTree']/div[@role = 'treeitem']/div[@role = 'group']")
    recursion(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)


def recursion(tree_root, wait, driver, root_node):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        root_node.add_child(new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('div')
        if div_element and div_element.get('aria-disabled') == 'true':
            wait_click_xpath_relative(wait, time_w = 0.5, element = item, xpath = "./div[@class = 'el-tree-node__content']/span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            new_tree_root = wait_return_subelement_relative(wait, time_w = 0.5, element= item, xpath = ".//div[@role = 'group']")
            recursion(new_tree_root, wait, driver, new_node)
        else:
            next



def synchronizing_job(wait, driver, input_node):
    #采集根组织信息
    root_text = 'root'
    input_node.set_value(root_text)
    #获取根节点结构体
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree objectTree']")
    recursion(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)


def recursion(tree_root, wait, driver, root_node):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        root_node.add_child(new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('div')
        if div_element and div_element.get('aria-disabled') == 'true':
            wait_click_xpath_relative(wait, time_w = 0.5, element = item, xpath = "./div[@class = 'el-tree-node__content']/span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            new_tree_root = wait_return_subelement_relative(wait, time_w = 0.5, element=item, xpath = ".//div[@role = 'group']")
            recursion(new_tree_root, wait, driver, new_node)
        else:
            next


def synchronizing_xueli(wait, driver, input_node):
    #采集根组织信息
    root_text = 'root'
    input_node.set_value(root_text)
    #获取根节点结构体
    tree_root = wait_return_subelement_absolute(wait, time_w = 0.5, xpath = "//div[@class = 'el-tree objectTree']")
    recursion(tree_root = tree_root, wait = wait, driver = driver, root_node = input_node)


def recursion(tree_root, wait, driver, root_node):
    org_items = []
    org_items = tree_root.find_elements(By.XPATH, ".//div[@role = 'treeitem']")
    for item in org_items:
        node_text = item.get_attribute('textContent')
        new_node =TreeNode(node_text)
        root_node.add_child(new_node)
        item_html = item.get_attribute('outerHTML')
        soup = BeautifulSoup(item_html, 'html.parser')
        div_element = soup.find('div')
        if div_element and div_element.get('aria-disabled') == 'true':
            wait_click_xpath_relative(wait, time_w = 0.5, element = item, xpath = "./div[@class = 'el-tree-node__content']/span[@class = 'el-tree-node__expand-icon el-icon-caret-right']")
            new_tree_root = wait_return_subelement_relative(wait, time_w = 0.5, element=item, xpath = ".//div[@role = 'group']")
            recursion(new_tree_root, wait, driver, new_node)
        else:
            next