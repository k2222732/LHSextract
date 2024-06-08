import base.waitclick as wc

#直接输入
def input_text(wait, driver, text):
    pass

#单日期选择输入
def select_date_single(wait, driver, date):
    pass

#单日期选择输入带分秒
def select_date_duration_sec(wait, driver, date, hms):
    pass

#区间选择日期
def select_date_duration(wait, driver, date_start, date_end):
    pass

#选择党组织
def select_party_org(wait, driver, org):
    pass

#选择背景
def select_background(wait, driver, background):
    pass

#选择工作岗位
def select_jobposition(wait, driver, job):
    pass

#单击通用按钮
def commen_button(wait, driver, xpath):
    pass

#上传图片
def upload_pic(pic_dir, input_xpath, button_xpath, wait, driver):
    '''
    pic_dir:图片路径
    input_xpath:输入位置的xpath
    button_xpath:保存按钮的xpath
    '''
    pass