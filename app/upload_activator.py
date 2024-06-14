import dev_func
from base.waitclick import wait_click_xpath
from selenium.webdriver.support.ui import WebDriverWait
import login
import os
import pandas as pd
from tkinter import messagebox
import base.rwconfig as conf
import base.write_entry as cursor
from base.boot import synchronizing_org, synchronizing_job, synchronizing_xueli
from base.mystruct import TreeNode

applicant_info = {}
org_tree = TreeNode
job_tree = TreeNode
education_tree = TreeNode

def main():
    global org_tree
    global job_tree
    global education_tree
    driver = login._main()
    wait = WebDriverWait(driver, 10, 0.5)
    dev_func.access_dev_database(driver, wait)
    dev_func.switch_role(wait, driver)
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "发展流程")])[1]')
    wait_click_xpath(wait, time_w = 0.5, xpath = '(//span[contains(text(), "添加")]')
    #载入党员信息表
    mfilepath = conf.pick_config_param(configfilepath = '..\\config.ini', str_class = '', str_param = '')
    global applicant_info
    applicant_info = load_to_cache(mfilepath)
    if applicant_info == False:
        return False
    
    #同步党组织树、职务树、学历树
    synchronizing_org(driver=driver, wait=wait, input_node=org_tree)
    synchronizing_job(driver=driver, wait=wait, input_node=job_tree)
    synchronizing_xueli(driver=driver, wait=wait, input_node=education_tree)
    #循环录入积极分子信息
    for person in mfilepath:
        step_1(driver, wait, data_dict=applicant_info, gmsfzhm=person['公民身份证号码'])
        step_2(driver, wait)
        step_3(driver, wait)
        step_4(driver, wait)
        step_5(driver, wait)
    

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
            qdrdjjfzrq = row['确定入党积极分子日期']
            gzdwjzw = row['工作单位及职务']
            jtzz = row['家庭住址']
            pylxr = row['培训联系人']
            data_dict[gmsfzhm] = {'序号': xh, '姓名': xm, '性别':xb, '公民身份证号码':gmsfzhm,'民族':mz,
                                  '出生日期':csny, '学历':xl, '申请入党日期':sqrdrq, '手机号码':sjhm, 
                                  '工作岗位':gzgw, '政治面貌':zzmm, '接受申请党组织':jssqdzz, '籍贯':jg ,
                                  '入团日期':rtrq, '参加工作日期':cjgzrq, '申请入党日期':sqrdrq, '确定入党积极分子日期':qdrdjjfzrq,
                                  '工作单位及职务':gzdwjzw, '家庭住址':jtzz, '培养联系人':pylxr, }
    return data_dict



def step_1(driver, wait, data_dict, gmsfzhm):
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
    #手机号码#(//input[@placeholder = '请输入手机号码'])[1]
    cursor.input_text(wait, driver, xpath="(//input[@placeholder = '请输入手机号码'])[1]", text=data_dict[gmsfzhm]['手机号码'])
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
    
    #接受申请党组织#(//input[@placeholder = '请输入接受申请党组织'])[1]
    #保存#(//span[contains(text(), '保存')])[1]


    #编辑#(//span[contains(text(), '编辑')])[2]
    #加号#//div[@class = 'el-upload el-upload--picture-card']
    #上传图片#//input[@type="file"]
    #保存#(//span[contains(text(), '保存')])[2]
    #提交#//span[contains(text(),'提交')]

    #警告#//div[@role = 'alert']
    #警告的内容#//div[@role = 'alert']//p[@class = 'el-message__content']
    pass


def step_2(driver, wait):
    #党组织派人谈话#(//a[contains(text(), '查看详情')])[2]
    #编辑#(//span[contains(text(), '编辑')])[1]
    #谈话人#(//input[@placeholder = '请输入内容'])[1]
    #选择日期#(//input[@placeholder = '选择日期'])[1]
    #上报#//span[contains(text(),'上报')]
    #审核#(//span[contains(text(),'审核')])[2]
    #审核通过#(//span[contains(text(),'审核通过')])[1]
    pass

def step_3(driver, wait):
    #推荐和确定入党积极分子#(//a[contains(text(), '查看详情')])[3]
    #编辑#(//span[contains(text(), '编辑')])[1]
    #党员推荐会议时间#(//input[@placeholder = '请选择日期'])[1]
    #保存#(//span[contains(text(), '保存')])[1]
    #上传双推会议纪录#(//input[@type="file"])[1]
    #编辑#(//span[contains(text(), '编辑')])[3]
    #时间#(//input[@placeholder = '请选择日期'])[1]
    #地点#(//input[@placeholder = '请输入内容'])[1]
    #主持人#(//input[@placeholder = '请输入内容'])[2]
    #会议类型#(//input[@placeholder = '请选择字段项'])[1]
    #党支部书记#(//input[@placeholder = '请输入内容'])[3]
    #列席组织员#(//input[@placeholder = '请输入内容'])[5]
    #会议研究意见#(//textarea[@placeholder = '请输入内容'])[1]
    #保存#(//span[contains(text(), '保存')])[3]


    #编辑#(//span[contains(text(), '编辑')])[4]
    #研究确定入党积极分子会议纪录#(//input[@type="file"])[3]
    #保存#(//span[contains(text(), '保存')])[3]

    #籍贯#(//input[@placeholder = '请选择字段项'])[1]
    #参加工作日期#(//input[@placeholder = '选择日期'])[2]
    #工作单位及职务#(//input[@placeholder = '请输入内容'])[1]
    #家庭住址#(//input[@placeholder = '请输入内容'])[2]
    #保存#(//span[contains(text(), '保存')])[4]

    #上报#//span[contains(text(),'上报')]

    #推荐和确定入党积极分子#(//a[contains(text(), '查看详情')])[3]
    #审核#(//span[contains(text(),'审核')])[2]
    #审核通过#(//span[contains(text(),'审核通过')])[1]
    pass


def step_4(driver, wait):
    #入党积极分子公示和备案#(//a[contains(text(), '查看详情')])[4]
    #编辑#(//span[contains(text(), '编辑')])[2]
    #公示起止日期#(//input[@placeholder = '开始日期'])[1]
    #公示结果#(//input[@placeholder = '请选择'])[1]
    #保存#(//span[contains(text(), '保存')])[2]

    #编辑#(//span[contains(text(), '编辑')])[3]
    #基层党委备案审查日期#(//input[@placeholder = '选择日期'])[1]
    #基层党委备案审查意见#(//input[@placeholder = '请选择字段项'])[1]
    #保存#(//span[contains(text(), '保存')])[2]
    #提交#//span[contains(text(),'提交')]
    pass


def step_5(driver, wait):
    #指定培养联系人#(//a[contains(text(), '查看详情')])[5]
    #添加#
    #姓名#
    #入党日期(现支部不需)#
    #学历(现支部不需)#
    #自何时起负责培养#
    #保存#
    #提交#

    pass





if __name__ == "__main__":
    main()