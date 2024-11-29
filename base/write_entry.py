from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import traceback
from dataclasses import dataclass
from base.mystruct import TreeNode
from base.waitclick import wait_click_xpath, wait_click_xpath_relative, wait_return_subelement_relative, wait_return_subelement_absolute, wait_return_subelement_absolute_notmust, wait_click_xpath_notmust
from base.waitclick import wait_return_subelement_absolute_notmust_visiable
import pandas as pd
import os
import time




from selenium.webdriver.support.ui import WebDriverWait



def click_button_until_specifyxpath_appear(wait, driver, specifyxpath, buttonxpath):
    while True:
        # 查找并点击按钮
        button = driver.find_element(By.XPATH, buttonxpath)
        button.click()
        try:
            # 等待目标元素出现
            wait.until(EC.presence_of_element_located((By.XPATH, specifyxpath)))
            break
        except:
            pass

@dataclass
class DATE:
    def __init__(self, year: int, month: int, day:int, hour: int, min:int, sec:int):
        self.year = year
        self.month = month
        self.day = day
        self.hour = hour
        self.min = min
        self.sec = sec
    year: int
    month: int
    day:int
    hour: int
    min:int
    sec:int


#直接输入
def input_text(wait, driver, xpath, text):
    try:
        input_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        input_element.clear()
        if isinstance(text, pd.Timestamp):
            input_element.send_keys(text._date_repr) 
        else:
            input_element.send_keys(text)
    except Exception:
        traceback.print_exc()

def input_text_v1(wait, driver, xpath, text, times):
    t = 1
    while 1:
        input_text(wait, driver, xpath=xpath, text=text)
        time.sleep(1)
        t = t+1
        if t >=times:
            break

#单日期选择输入
def select_date_single(wait, driver, xpath, date):
    try:
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        date_input.clear()
        date_str = f"{date.year:04d}-{date.month:02d}-{date.day:02d}"
        # 将格式化的日期字符串输入到输入框中
        date_input.send_keys(date_str)
    except Exception:
        traceback.print_exc()

#单日期选择输入带分秒
def select_date_duration_sec(wait, driver, xpath, date:DATE):
    try:
        date_input = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        date_input.clear()
        date_str = f"{date.year:04d}-{date.month:02d}-{date.day:02d} {date.hour:02d}:{date.min:02d}"
        date_input.send_keys(date_str)
    except Exception:
        traceback.print_exc()
    pass

#区间选择日期
def select_date_duration(wait, driver, date_start:DATE, date_end:DATE, start_xpath, end_xpath):
    try:
        start_input = wait.until(EC.presence_of_element_located((By.XPATH, start_xpath)))
        start_input.clear()
        start_input.send_keys(f"{date_start.year:04d}-{date_start.month:02d}-{date_start.day:02d}")

        end_input = wait.until(EC.presence_of_element_located((By.XPATH, end_xpath)))
        end_input.clear()
        end_input.send_keys(f"{date_end.year:04d}-{date_end.month:02d}-{date_end.day:02d}")
    except Exception:
        traceback.print_exc()



#点选列表（单个）
def select_background(wait, driver, element_xpath, label='span', value='汉族'):
    try:
        path = f'{element_xpath}//{label}[contains(text(), "{value}")]'
        wait_click_xpath(wait, time_w = 0.5, xpath = path)
    except Exception:
        traceback.print_exc()

#点选列表（单个）
def select_background_v1(wait, driver, element_xpath, label='span', value='汉族'):
    try:
        path = f'{element_xpath}//{label}[contains(text(), "{value}")]/..'
        wait_click_xpath(wait, time_w = 0.5, xpath = path)
    except Exception:
        traceback.print_exc()





#点选列表（树）
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
            wait_click_xpath_relative(wait, time_w = 0.5, element = current_element, xpath = f".//{label_1}[contains(text(), '{node_value}')]/../preceding-sibling::span")
                                                                                                                    #(//div[@role = 'treeitem'])[1]//span[contains(text(), '中国共产党山东汶上经济开发区工作委员会')]/../../following-sibling::div[@role = 'group']
            current_element = wait_return_subelement_relative(time_w = 0.5, element = current_element, xpath = f".//{label_1}[contains(text(), '{node_value}')]/../../following-sibling::div[@role = 'group']")
    #如果node_value是列表中的最后一个元素，则点击元素体     
        else:                                                #测试：  (//div[@role = 'treeitem'])[1]//span[contains(text(), '中国共产党山东汶上经济开发区工作委员会')]/../../following-sibling::div[@role = 'group']/.//span[contains(text(), '中国共产党新风光电子科技股份有限公司委员会')]/../preceding-sibling::span
            wait_click_xpath_relative(wait, time_w = 0.5, element = current_element, xpath = f".//{label_2}[contains(text(), '{node_value}')]")                                                                                                     #测试： (//div[@role = 'treeitem'])[1]//span[contains(text(), '中国共产党山东汶上经济开发区工作委员会')]/../../following-sibling::div[@role = 'group']//span[contains(text(), '中国共产党新风光电子科技股份有限公司第一支部委员会')]


#单击通用按钮
def commen_button(wait, driver, xpath):
    
    try:
        wait_click_xpath(wait, time_w = 0.5, xpath = xpath)
    except Exception:
        traceback.print_exc()


def commen_button_w(wait, xpath):
    try:
        wait_click_xpath(wait, time_w = 0.5, xpath = xpath)
    except Exception:
        traceback.print_exc()


#单击通用按钮非一定
def commen_button_notmust(wait, driver, xpath):
    
    try:
        wait_click_xpath_notmust(wait, time_w = 0.5, xpath = xpath)
    except Exception:
        traceback.print_exc()


#点击按钮直到某个xpath消失
def click_button_until_specifyxpath_disappear(wait, driver, specifyxpath, buttonxpath, time_w, times):
    while 1:
        yanzheng = wait_return_subelement_absolute_notmust(wait=wait, time_w=time_w, xpath=specifyxpath, times=times)
        if yanzheng is not None:
            try:
                commen_button_notmust(wait, driver, xpath=buttonxpath)
                time.sleep(1)
            except:
                pass
        elif yanzheng is None: 
            break

def click_button_until_specifyxpath_appear(wait, driver, specifyxpath, buttonxpath):
    while True:
        # 查找并点击按钮
        try:
            button = wait.until(EC.element_to_be_clickable((By.XPATH, buttonxpath)))
            button.click()
        except:
            pass
        try:
            # 等待目标元素出现
            wait.until(EC.presence_of_element_located((By.XPATH, specifyxpath)))
            break
        except:
            pass


def click_button_until_specifyxpath_appear_except(wait, driver, specifyxpath, buttonxpath, time_w, times):
    
    i = 1
    while True:
        if i >= times:
            raise TimeoutException("主动抛出超时异常")
        # 查找并点击按钮
        try:
            button = WebDriverWait(driver, 4).until(EC.element_to_be_clickable((By.XPATH, buttonxpath)))
            button.click()
        except:
            pass
        try:
            # 等待目标元素出现
            WebDriverWait(driver, time_w).until(EC.presence_of_element_located((By.XPATH, specifyxpath)))
            break
        except:
            i = i +1
            

#等待某个xpath出现，然后再消失
def wait_specifyxpath_appear_disappear(wait, driver, xpath):
    wait.until(EC.visibility_of_element_located(By.XPATH, xpath))
    while 1:
        t = wait.until(EC.visibility_of_element_located(By.XPATH, xpath))
        if t is not None:
            time.sleep(0.5)
        elif t is None:
            break




#上传图片
def upload_pic(pic_dir, input_xpath, wait, driver):
    '''
    pic_dir:图片路径
    input_xpath:输入位置的xpath
    '''
    try:
        files = [os.path.join(pic_dir, f) for f in os.listdir(pic_dir) if os.path.isfile(os.path.join(pic_dir, f))]
        file_input = wait.until(EC.presence_of_element_located((By.XPATH, input_xpath)))
        for file in files:
            file_input.send_keys(file)
    except Exception:
        traceback.print_exc()

