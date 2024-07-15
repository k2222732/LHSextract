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





def main():
    driver = login._main()
    wait = WebDriverWait(driver, 2, 0.5)
    middle.switch_role.access_e_shandong(driver, wait)
    middle.switch_role.switch_role(wait)
    
    input("")


if __name__ == "__main__":
    main()