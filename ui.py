import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import os
import subprocess
import configparser
import threading
import member
import org
import dev
import member_func
import org_func
import dev_func
import login
import vip_struct
import ui_button_func

role_name_mem = ""
role_name_org = ""
role_name_dev = ""

##############
class MainApp:
    def __init__(self, root2):
        self.mem_thread = None
        self.org_thread = None
        self.dev_thread = None

        self.config_file = 'config.ini'
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file)
        self.root2 = root2
        self.root2.title("灯塔填表小助手")
        self.tabControl = ttk.Notebook(self.root2)
        #self.root2.bind('q', self.delete_password_field)
        
        

        # 选项卡1
        self.tab1 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab1, text="基本设置")
        self.setup_tab1()
        self.privious_tab = self.tab1
        self.tabControl.bind("<<NotebookTabChanged>>", self.on_tab_change)

        # 选项卡2
        self.tab2 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab2, text="仅登录")
        #self.setup_tab2()
        


        # 选项卡3
        self.tab3 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab3, text="数据同步")
        #self.setup_tab3()
        

        # 选项卡4
        self.tab4 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab4, text="主题党日上传")
        #self.setup_tab4()
        
        

        # 选项卡5
        self.tab5 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab5, text="发展党员材料上传")
        #self.setup_tab5()


        # 选项卡6
        self.tab6 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab6, text="自动填表")
        #self.setup_tab6()


        # 选项卡7
        self.tab7 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab7, text="账号信息")
        #self.setup_tab7()


        #self.tabControl.pack(expand=1, fill="both")
        self.tabControl.grid(row=0, column=0, sticky="nsew")
        


    def on_tab_change(self, event = None):
        n = 1
        current_tab = self.tabControl.tab(self.tabControl.select(), "text")
        privious_tab = self.tabControl.tab(self.privious_tab, "text")
        if current_tab == '基本设置':
            num_columns, num_rows = self.privious_tab.grid_size()
            while n <= num_rows-1:
                slaves = self.privious_tab.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
    

            self.setup_tab1()
            self.load_config()
            self.privious_tab = self.tab1
            
        elif current_tab == '仅登录':
            num_columns, num_rows = self.privious_tab.grid_size()
            while n <= num_rows-1:
                slaves = self.privious_tab.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            


            self.setup_tab2()
            self.privious_tab = self.tab2
            
        elif current_tab == '数据同步':
            num_columns, num_rows = self.privious_tab.grid_size()
            while n <= num_rows-1:
                slaves = self.privious_tab.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1

            self.setup_tab3()
            self.privious_tab = self.tab3

        elif current_tab == '主题党日上传':
            num_columns, num_rows = self.privious_tab.grid_size()
            while n <= num_rows-1:
                slaves = self.privious_tab.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            

            self.setup_tab4()
            self.privious_tab = self.tab4

        elif current_tab == '发展党员材料上传':
            num_columns, num_rows = self.privious_tab.grid_size()
            while n <= num_rows-1:
                slaves = self.privious_tab.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            

            self.setup_tab5()
            self.privious_tab = self.tab5

        elif current_tab == '账号信息':
            num_columns, num_rows = self.privious_tab.grid_size()
            while n <= num_rows-1:
                slaves = self.privious_tab.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            

            self.setup_tab7()
            self.privious_tab = self.tab7
    
        


    def load_config(self):
        if os.path.exists(self.config_file):
            self.explore_driver_path.insert(0, self.config.get('Paths_driver', 'explore_driver_path', fallback=''))
            self.explore_path.insert(0, self.config.get('Paths_explore', 'explore_path', fallback=''))
            self.lhaccount.insert(0, self.config.get('Account', 'account', fallback=''))
            self.lhpassword.insert(0, self.config.get('Password', 'password', fallback=''))
        else:
            pass


    def setup_tab1(self):
        # 创建 Label 和 Entry 用于显示和选择目录
        ttk.Label(self.tab1, text="浏览器驱动位置：").grid(row=0, column=0, sticky="w", pady=5)
        self.explore_driver_path = ttk.Entry(self.tab1, state="normal")
        self.explore_driver_path.grid(row=0, column=1, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        ttk.Button(self.tab1, text="选择目录", command=self.choose_explore_driver_path).grid(row=0, column=3, pady=5, padx=(10, 5))
        ttk.Button(self.tab1, text="保存", command=self.save_explore_driver_path).grid(row=0, column=4, pady=5, padx=(5, 10))

        ttk.Label(self.tab1, text="浏览器位置：").grid(row=1, column=0, sticky="w", pady=5)
        self.explore_path = ttk.Entry(self.tab1, state="normal")
        self.explore_path.grid(row=1, column=1, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        ttk.Button(self.tab1, text="选择目录", command=self.choose_explore_path).grid(row=1, column=3, pady=5, padx=(10, 5))
        ttk.Button(self.tab1, text="保存", command=self.save_explore_path).grid(row=1, column=4, pady=5, padx=(5, 10))

        ttk.Label(self.tab1, text="灯塔账号：").grid(row=2, column=0, sticky="w", pady=5)
        self.lhaccount = ttk.Entry(self.tab1, state="normal")
        self.lhaccount.grid(row=2, column=1, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        ttk.Button(self.tab1, text="保存", command=self.save_lhaccount).grid(row=2, column=3, pady=5, padx=(10, 10))
        
        ttk.Label(self.tab1, text="灯塔密码：").grid(row=3, column=0, sticky="w", pady=5)
        self.lhpassword = ttk.Entry(self.tab1, state="normal")
        self.lhpassword.grid(row=3, column=1, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        ttk.Button(self.tab1, text="保存", command=self.save_lhpassword).grid(row=3, column=3, pady=5, padx=(10, 10))



    def setup_tab2(self):
        ttk.Button(self.tab2, text="一键登录", command=self.open_login).grid(row=0, column=1, pady=5, padx=(10, 5))
        


    def setup_tab3(self):
        ttk.Button(self.tab3, text="同步党员信息", command=self.open_mem_thread).grid(row=0, column=1, pady=5, padx=(10, 5))
        ttk.Button(self.tab3, text="选择目录", command=self.choose_mem_info_path).grid(row=0, column=2, pady=5, padx=(10, 5))
        self.role_mem = ttk.Entry(self.tab3, state="normal")
        self.role_mem.grid(row=0, column=3, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        self.role_mem.insert(0, self.config.get('role_mem_name', 'name_role_mem', fallback=''))
        ttk.Button(self.tab3, text="停止", command=self.close_thread_mem).grid(row=0, column=5, pady=5, padx=(10, 5))

        ttk.Button(self.tab3, text="同步党组织信息", command=self.open_org_thread).grid(row=1, column=1, pady=5, padx=(10, 5))
        ttk.Button(self.tab3, text="选择目录", command=self.choose_org_info_path).grid(row=1, column=2, pady=5, padx=(10, 5))
        self.role_org = ttk.Entry(self.tab3, state="normal")
        self.role_org.grid(row=1, column=3, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        self.role_org.insert(0, self.config.get('role_org_name', 'name_role_org', fallback=''))
        ttk.Button(self.tab3, text="停止", command=self.close_thread_org).grid(row=1, column=5, pady=5, padx=(10, 5))

        ttk.Button(self.tab3, text="同步发展纪实信息", command=self.open_dev_thread).grid(row=2, column=1, pady=5, padx=(10, 5))
        ttk.Button(self.tab3, text="选择目录", command=self.choose_dev_info_path).grid(row=2, column=2, pady=5, padx=(10, 5))
        self.role_dev = ttk.Entry(self.tab3, state="normal")
        self.role_dev.grid(row=2, column=3, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        self.role_dev.insert(0, self.config.get('role_dev_name', 'name_role_dev', fallback=''))
        ttk.Button(self.tab3, text="停止", command=self.close_thread_dev).grid(row=2, column=5, pady=5, padx=(10, 5))
        
        


    def setup_tab4(self):
        self.combo_box = ttk.Combobox(self.tab4, values=["单支部上传", "多支部批量上传"], width=18)
        self.combo_box.grid(row=0, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.combo_box.bind("<<ComboboxSelected>>", self.show_layout)
        ##单支部上传控件创建##
        self.role_to_up_load_ztdr = ttk.Entry(self.tab4, state="normal", width=20)
        self.role_to_up_load_ztdr.grid(row=1, column=0, columnspan=2, sticky="w", pady=5, padx=(10, 0))
        self.label_instruction = tk.Label(self.tab4, text="请输入角色名称")
        self.label_instruction.grid(row=2, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.var = tk.StringVar(master=self.tab4)
        self.radio_button1 = tk.Radiobutton(self.tab4, text="主题党日", variable=self.var, value="选项1")
        self.radio_button2 = tk.Radiobutton(self.tab4, text="灯塔大课堂", variable=self.var, value="选项2")
        self.radio_button3 = tk.Radiobutton(self.tab4, text="主题党日+灯塔大课堂", variable=self.var, value="选项3")
        self.radio_button4 = tk.Radiobutton(self.tab4, text="组织生活会", variable=self.var, value="选项4")
        self.radio_button5 = tk.Radiobutton(self.tab4, text="民主生活会", variable=self.var, value="选项5")
        self.var0 = tk.StringVar(master=self.tab4)
        self.radio_button1_way = tk.Radiobutton(self.tab4, text="以文本形式上传", variable=self.var0, value="文本")
        self.label_instruction = tk.Label(self.tab4, text="在此输入文字版会议纪录：")
        self.role_to_up_load_ztdr = tk.Text(self.tab4, state="normal", width=20, height = 13)
        self.button_generate = ttk.Button(self.tab4, text="一键生成会议纪录", command=self.generate_txt)
        self.radio_button2_way = tk.Radiobutton(self.tab4, text="以照片形式上传", variable=self.var0, value="照片")
        self.pictures_dir = ttk.Label(self.tab4, text="请选择图片所在目录")
        self.pictures_path = ttk.Entry(self.tab4, state="normal", width = 15)
        self.button_choose_dir = ttk.Button(self.tab4, text="选择目录", command=self.choose_pictures_path, width = 8)
        self.button_start = ttk.Button(self.tab4, text="点击启动", command=self.start_upload_solo)
        #单支部布局
        self.radio_button1.grid(row=3, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.radio_button2.grid(row=4, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.radio_button3.grid(row=5, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.radio_button4.grid(row=6, column=0, columnspan=2, sticky="w", pady=3, padx=(10, 0))
        self.radio_button5.grid(row=7, column=0, columnspan=2, sticky="w", pady=3, padx=(10, 0))
        self.radio_button1_way.grid(row=0, column=2, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.label_instruction.grid(row=1, column=2, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.role_to_up_load_ztdr.grid(row=2, column=2, columnspan=1, rowspan=5, sticky="w", pady=0, padx=(10, 0))
        self.button_generate.grid(row=7, column=2, rowspan=2, pady=0, padx=(10, 0))
        self.radio_button2_way.grid(row=0, column=3, columnspan=1, sticky="w", pady=0, padx=(10, 0))
        self.pictures_dir.grid(row=1, column=3, sticky="w", pady=0, padx=(15, 0))
        self.pictures_path.grid(row=2, column=3, columnspan=2, sticky="w", pady=5, padx=(15, 0))
        self.pictures_path.insert(0, self.config.get('tab4_path1', 'ztdr_pic_path', fallback=''))
        self.button_choose_dir.grid(row=3, column=3, pady=0, padx=(15, 0), sticky="w")
        self.button_start.grid(row=4, column=3, rowspan=1, pady=0, padx=(15, 0), sticky="w")
        ##多支部批量上传控件创建##
        
        
        
        


    def show_layout(self, event = None):
        selected_item = self.combo_box.get()
        n = 1
        if selected_item == "单支部上传":
            num_columns, num_rows = self.tab4.grid_size()
            slave02 = self.tab4.grid_slaves(row=0, column=2)
            slave03 = self.tab4.grid_slaves(row=0, column=3)
            for slave in slave02:
                slave.destroy()
            for slave in slave03:
                slave.destroy()
            while n <= num_rows:
                slaves = self.tab4.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
        #单支部布局
            self.role_to_up_load_ztdr = ttk.Entry(self.tab4, state="normal", width=20)
            self.role_to_up_load_ztdr.grid(row=1, column=0, columnspan=2, sticky="w", pady=5, padx=(10, 0))
            self.label_instruction = tk.Label(self.tab4, text="请输入角色名称")
            self.label_instruction.grid(row=2, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            self.var = tk.StringVar(master=self.tab4)
            self.radio_button1 = tk.Radiobutton(self.tab4, text="主题党日", variable=self.var, value="选项1")
            self.radio_button2 = tk.Radiobutton(self.tab4, text="灯塔大课堂", variable=self.var, value="选项2")
            self.radio_button3 = tk.Radiobutton(self.tab4, text="主题党日+灯塔大课堂", variable=self.var, value="选项3")
            self.radio_button4 = tk.Radiobutton(self.tab4, text="组织生活会", variable=self.var, value="选项4")
            self.radio_button5 = tk.Radiobutton(self.tab4, text="民主生活会", variable=self.var, value="选项5")
            self.var0 = tk.StringVar(master=self.tab4)
            self.radio_button1_way = tk.Radiobutton(self.tab4, text="以文本形式上传", variable=self.var0, value="文本")
            self.label_instruction = tk.Label(self.tab4, text="在此输入文字版会议纪录：")
            self.role_to_up_load_ztdr = tk.Text(self.tab4, state="normal", width=20, height = 13)
            self.button_generate = ttk.Button(self.tab4, text="一键生成会议纪录", command=self.generate_txt)
            self.radio_button2_way = tk.Radiobutton(self.tab4, text="以照片形式上传", variable=self.var0, value="照片")
            self.pictures_dir = ttk.Label(self.tab4, text="请选择图片所在目录")
            self.pictures_path = ttk.Entry(self.tab4, state="normal", width = 15)
            self.button_choose_dir = ttk.Button(self.tab4, text="选择目录", command=self.choose_pictures_path, width = 8)
            self.button_start = ttk.Button(self.tab4, text="点击启动", command=self.start_upload_solo)
            #单支部布局
            self.radio_button1.grid(row=3, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            self.radio_button2.grid(row=4, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            self.radio_button3.grid(row=5, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            self.radio_button4.grid(row=6, column=0, columnspan=2, sticky="w", pady=3, padx=(10, 0))
            self.radio_button5.grid(row=7, column=0, columnspan=2, sticky="w", pady=3, padx=(10, 0))
            self.radio_button1_way.grid(row=0, column=2, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            self.label_instruction.grid(row=1, column=2, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            self.role_to_up_load_ztdr.grid(row=2, column=2, columnspan=1, rowspan=5, sticky="w", pady=0, padx=(10, 0))
            self.button_generate.grid(row=7, column=2, rowspan=2, pady=0, padx=(10, 0))
            self.radio_button2_way.grid(row=0, column=3, columnspan=1, sticky="w", pady=0, padx=(10, 0))
            self.pictures_dir.grid(row=1, column=3, sticky="w", pady=0, padx=(15, 0))
            self.pictures_path.grid(row=2, column=3, columnspan=2, sticky="w", pady=5, padx=(15, 0))
            self.pictures_path.insert(0, self.config.get('tab4_path1', 'ztdr_pic_path', fallback=''))
            self.button_choose_dir.grid(row=3, column=3, pady=0, padx=(15, 0), sticky="w")
            self.button_start.grid(row=4, column=3, rowspan=1, pady=0, padx=(15, 0), sticky="w")
        #多支部布局
        elif selected_item == "多支部批量上传":
            num_columns, num_rows = self.tab4.grid_size()
            slave02 = self.tab4.grid_slaves(row=0, column=2)
            slave03 = self.tab4.grid_slaves(row=0, column=3)
            for slave in slave02:
                slave.destroy()
            for slave in slave03:
                slave.destroy()
            while n <= num_rows:
                slaves = self.tab4.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            self.var = tk.StringVar(master=self.tab4)
            self.radio_button1 = tk.Radiobutton(self.tab4, text="主题党日", variable=self.var, value="选项1")          
            self.radio_button2 = tk.Radiobutton(self.tab4, text="灯塔大课堂", variable=self.var, value="选项2")           
            self.radio_button3 = tk.Radiobutton(self.tab4, text="主题党日+灯塔大课堂", variable=self.var, value="选项3")    
            self.label_instruction = tk.Label(self.tab4, text="在此输入文字版会议纪录：")        
            self.role_to_up_load_ztdr = tk.Text(self.tab4, state="normal", width=20, height = 13)         
            self.click_generate = ttk.Button(self.tab4, text="一键生成会议纪录", command=self.generate_txt)           
            self.click_start = ttk.Button(self.tab4, text="点击启动", command=self.start_upload_solo)            
            self.label_instruction2 = tk.Label(self.tab4, text="说明：批量上传将自动检测未上\n传本月主题党日和灯塔大课堂\n的支部，并自动上传。")
            self.radio_button1.grid(row=1, column=0, columnspan=1, sticky="w", pady=0, padx=(10, 0))
            self.radio_button2.grid(row=2, column=0, columnspan=1, sticky="w", pady=(0, 10), padx=(10, 0))
            self.radio_button3.grid(row=3, column=0, columnspan=1, sticky="w", pady=0, padx=(10, 0))
            self.label_instruction.grid(row=0, column=2, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            self.role_to_up_load_ztdr.grid(row=1, column=2, columnspan=1, rowspan=3, sticky="w", pady=0, padx=(10, 0))
            self.click_generate.grid(row=6, column=2, rowspan=2, pady=0, padx=(10, 0))
            self.click_start.grid(row=0, column=3, rowspan=1, pady=(10,0), padx=(15, 0), sticky="e")
            self.label_instruction2.grid(row=1, column=3, columnspan=2, sticky="w", pady=0, padx=(10, 0))



    def start_upload_solo(self):
        pass

    def generate_txt(self):
        pass

    def choose_pictures_path(self):
        path = filedialog.askdirectory()
        self.pictures_path.delete(0, tk.END)
        self.pictures_path.insert(0, path)
        self.config['tab4_path1'] = {
                    'ztdr_pic_path': self.pictures_path.get()
                }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)


    def setup_tab5(self):
        self.label_instruction1 = tk.Label(self.tab5, text="请选择需上传到哪个阶段")  
        self.label_instruction1.grid(row=0, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.combo_box_dev = ttk.Combobox(self.tab5, values=["1、上传到积极分子阶段", "2、上传到发展对象阶段", "3、上传到预备党员阶段"], width=18)
        self.combo_box_dev.grid(row=1, column=0, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        self.combo_box_dev.bind("<<ComboboxSelected>>", self.show_layout_upload_dev)

        self.tab5_frame1 = ttk.Frame(self.tab5)
        self.tab5_frame1.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        self.label_instruction2 = tk.Label(self.tab5_frame1, text="1、第一步:创建模板")  
        self.label_instruction2.grid(row=0, column=0, columnspan=1, sticky="w", pady=0, padx=(10, 0))
        self.excel_template_path = ttk.Entry(self.tab5_frame1, state="normal")
        self.excel_template_path.grid(row=1, column=0)
        self.excel_template_path.insert(0, self.config.get('tab5_path1', 'activist_info_path', fallback=''))
        self.tab5_button_choose_dir_1 = ttk.Button(self.tab5_frame1, text="选择目录", command=self.tab5_button_func_1, width = 8)
        self.tab5_button_choose_dir_1.grid(row=1, column=1, padx=(10, 0))
        self.tab5_button_1 = ttk.Button(self.tab5_frame1, text="点击生成积极分子信息采集模板", command=self.tab5_button_func_2, width = 31)
        self.tab5_button_1.grid(row=2, column=0, columnspan=2, sticky="w")

        self.tab5_frame2 = ttk.Frame(self.tab5)
        self.tab5_frame2.grid(row=4, column=0, columnspan=2, sticky="w", pady=5, padx=(10, 0))
        self.label_instruction3 = tk.Label(self.tab5_frame2, text="2、第二步:请打开模板，并按行输入积极分子个人信息。\n     请保证该文件夹下只有一个模板文件", justify="left")  
        self.label_instruction3.grid(row=0, column=0, columnspan=1, sticky="w", padx=(10, 0))
        self.tab5_button_2 = ttk.Button(self.tab5_frame2, text="打开模板目录", command=self.tab5_button_func_3, width = 31)
        self.tab5_button_2.grid(row=1, column=0, columnspan=2, sticky="w")

        self.tab5_frame3 = ttk.Frame(self.tab5)
        self.tab5_frame3.grid(row=5, column=0, sticky="w", pady=5, padx=(10, 0))
        self.label_instruction4 = tk.Label(self.tab5_frame3, text="3、第三步:选择积极分子纸质材料照片的文件夹路径。请\n严格按格式命名文件夹、放置材料图片，小程序会自动识\n别身份证号码与第一步excel模板条目对应", justify="left")  
        self.label_instruction4.grid(row=0, column=0, columnspan=2, sticky="w", padx=(10, 0))
        self.tab5_entry_2 = ttk.Entry(self.tab5_frame3, state="normal")
        self.tab5_entry_2.grid(row=1, column=0, sticky="w", columnspan=2)
        self.tab5_entry_2.insert(0, self.config.get('tab5_path2', 'activist_pic_path', fallback=''))
        self.tab5_button_3 = ttk.Button(self.tab5_frame3, text="选择目录", command=self.tab5_button_func_4, width = 8)
        self.tab5_button_3.grid(row=1, column=1, sticky="w", padx=(65, 0))

        self.tab5_frame4 = ttk.Frame(self.tab5)
        self.tab5_frame4.grid(row=6, column=0, sticky="w", pady=5, padx=(10, 0))
        self.label_instruction5 = tk.Label(self.tab5_frame4, text="4、第四步:开始上传", justify="left")  
        self.label_instruction5.grid(row=0, column=0, columnspan=2, sticky="w", padx=(10, 0))
        self.tab5_button_4 = ttk.Button(self.tab5_frame4, text="开始上传", command=self.tab5_button_func_5, width = 8)
        self.tab5_button_4.grid(row=1, column=0, sticky="w", padx=(10, 0))
        

    def tab5_button_func_1(self):
        path = filedialog.askdirectory()
        self.excel_template_path.delete(0, tk.END)
        self.excel_template_path.insert(0, path)
        self.config['tab5_path1'] = {
                    'activist_info_path': self.excel_template_path.get()
                }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)
        pass


    def tab5_button_func_2(self):
        

        pass


    def tab5_button_func_3(self):

        pass


    def tab5_button_func_4(self):
        path = filedialog.askdirectory()
        self.tab5_entry_2.delete(0, tk.END)
        self.tab5_entry_2.insert(0, path)
        self.config['tab5_path2'] = {
                    'activist_pic_path': self.tab5_entry_2.get()
                }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)
        pass

    def tab5_button_func_5(self):
        pass

    def show_layout_upload_dev(self, event = None):
        selected_item = self.combo_box_dev.get()
        n = 2
        if selected_item == "1、上传到积极分子阶段":
            num_columns, num_rows = self.tab5.grid_size()
            slave02 = self.tab5.grid_slaves(row=1, column=2)
            slave03 = self.tab5.grid_slaves(row=1, column=3)
            for slave in slave02:
                slave.destroy()
            for slave in slave03:
                slave.destroy()
            while n <= num_rows:
                slaves = self.tab5.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            pass
            self.tab5_frame1 = ttk.Frame(self.tab5)
            self.tab5_frame1.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
            self.label_instruction2 = tk.Label(self.tab5_frame1, text="1、第一步:创建模板")  
            self.label_instruction2.grid(row=0, column=0, columnspan=1, sticky="w", pady=0, padx=(10, 0))
            self.excel_template_path = ttk.Entry(self.tab5_frame1, state="normal")
            self.excel_template_path.grid(row=1, column=0)
            self.excel_template_path.insert(0, self.config.get('tab5_path1', 'activist_info_path', fallback=''))
            self.tab5_button_choose_dir_1 = ttk.Button(self.tab5_frame1, text="选择目录", command=self.tab5_button_func_1, width = 8)
            self.tab5_button_choose_dir_1.grid(row=1, column=1, padx=(10, 0))
            self.tab5_button_1 = ttk.Button(self.tab5_frame1, text="点击生成积极分子信息采集模板", command=ui_button_func.tab5_button_func_2, width = 31)
            self.tab5_button_1.grid(row=2, column=0, columnspan=2, sticky="w")

            self.tab5_frame2 = ttk.Frame(self.tab5)
            self.tab5_frame2.grid(row=4, column=0, columnspan=2, sticky="w", pady=5, padx=(10, 0))
            self.label_instruction3 = tk.Label(self.tab5_frame2, text="2、第二步:请打开模板，并按行输入积极分子个人信息。\n     请保证该文件夹下只有一个模板文件", justify="left")  
            self.label_instruction3.grid(row=0, column=0, columnspan=1, sticky="w", padx=(10, 0))
            self.tab5_button_2 = ttk.Button(self.tab5_frame2, text="打开模板目录", command=self.tab5_button_func_3, width = 31)
            self.tab5_button_2.grid(row=1, column=0, columnspan=2, sticky="w")

            self.tab5_frame3 = ttk.Frame(self.tab5)
            self.tab5_frame3.grid(row=5, column=0, sticky="w", pady=5, padx=(10, 0))
            self.label_instruction4 = tk.Label(self.tab5_frame3, text="3、第三步:选择积极分子纸质材料照片的文件夹路径。请\n严格按格式命名文件夹、放置材料图片，小程序会自动识\n别身份证号码与第一步excel模板条目对应", justify="left")  
            self.label_instruction4.grid(row=0, column=0, columnspan=2, sticky="w", padx=(10, 0))
            self.tab5_entry_2 = ttk.Entry(self.tab5_frame3, state="normal")
            self.tab5_entry_2.grid(row=1, column=0, sticky="w", columnspan=2)
            self.tab5_entry_2.insert(0, self.config.get('tab5_path2', 'activist_pic_path', fallback=''))
            self.tab5_button_3 = ttk.Button(self.tab5_frame3, text="选择目录", command=self.tab5_button_func_4, width = 8)
            self.tab5_button_3.grid(row=1, column=1, sticky="w", padx=(65, 0))

            self.tab5_frame4 = ttk.Frame(self.tab5)
            self.tab5_frame4.grid(row=6, column=0, sticky="w", pady=5, padx=(10, 0))
            self.label_instruction5 = tk.Label(self.tab5_frame4, text="4、第四步:开始上传", justify="left")  
            self.label_instruction5.grid(row=0, column=0, columnspan=2, sticky="w", padx=(10, 0))
            self.tab5_button_4 = ttk.Button(self.tab5_frame4, text="开始上传", command=self.tab5_button_func_5, width = 8)
            self.tab5_button_4.grid(row=1, column=0, sticky="w", padx=(10, 0))
        elif selected_item == "2、上传到发展对象阶段":
            num_columns, num_rows = self.tab5.grid_size()
            slave02 = self.tab5.grid_slaves(row=1, column=2)
            slave03 = self.tab5.grid_slaves(row=1, column=3)
            for slave in slave02:
                slave.destroy()
            for slave in slave03:
                slave.destroy()
            while n <= num_rows:
                slaves = self.tab5.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            pass

            self.label_instruction3 = tk.Label(self.tab5, text="敬请期待！")  
            self.label_instruction3.grid(row=1, column=2, columnspan=2, sticky="w", pady=0, padx=(10, 0))
            
        elif selected_item == "3、上传到预备党员阶段":
            num_columns, num_rows = self.tab5.grid_size()
            slave02 = self.tab5.grid_slaves(row=1, column=2)
            slave03 = self.tab5.grid_slaves(row=1, column=3)
            for slave in slave02:
                slave.destroy()
            for slave in slave03:
                slave.destroy()
            while n <= num_rows:
                slaves = self.tab5.grid_slaves(row=n)
                for slave in slaves:
                    slave.destroy()
                n = n + 1
            pass
        
            self.label_instruction4 = tk.Label(self.tab5, text="敬请期待！")  
            self.label_instruction4.grid(row=1, column=2, columnspan=2, sticky="w", pady=0, padx=(10, 0))
        pass

    
    def setup_tab6(self):
        pass

    def setup_tab7(self):
        self.vip_info_frame = ttk.Frame(self.tab7)
        self.vip_info_frame.grid(row=0, column=0, sticky="w")
        self.vip_label = tk.Label(self.vip_info_frame, text="会员信息：")
        self.vip_label.grid(row=0, column=0, sticky="w")
        self.vip_start_time = tk.Label(self.vip_info_frame, text="开通时间：")
        self.vip_start_time.grid(row=1, column=0, sticky="w")
        self.vip_type = tk.Label(self.vip_info_frame, text="会员类型：")
        self.vip_type.grid(row=2, column=0, sticky="w")
        self.vip_deadline = tk.Label(self.vip_info_frame, text="到期时间：")
        self.vip_deadline.grid(row=3, column=0, sticky="w")

        global vip_info

        self.is_vip = tk.Label(self.vip_info_frame, text=vip_info.vip)
        self.is_vip.grid(row=0, column=1, sticky="w")
        self.vip_start_time_value = tk.Label(self.vip_info_frame, text=vip_info.vip_start_time)
        self.vip_start_time_value.grid(row=1, column=1, sticky="w")
        self.vip_type_value = tk.Label(self.vip_info_frame, text=vip_info.type)
        self.vip_type_value.grid(row=2, column=1, sticky="w")
        self.vip_deadline_value = tk.Label(self.vip_info_frame, text=vip_info.deadline)
        self.vip_deadline_value.grid(row=3, column=1, sticky="w")


        self.vip_active = ttk.Frame(self.tab7)
        self.vip_active.grid(row=5, column=0, pady=10)
        self.activation_code = tk.Label(self.vip_active, text="输入激活码：")
        self.activation_code.grid(row=0, column=0, sticky="w")
        self.activation_input = ttk.Entry(self.vip_active)
        self.activation_input.grid(row=1, column=0, sticky="w", padx=3, pady=3)
        self.button_active = ttk.Button(self.vip_active, text="激活/续期", command = self.vip_active)
        self.button_active.grid(row=1, column=2, sticky="w", padx=3, pady=3)


        self.buy_activation_code = tk.Label(self.vip_active, text="购买激活码", cursor="hand2", foreground="blue")
        self.buy_activation_code.bind("<Button-1>", self.open_url)
        self.buy_activation_code.grid(row=2, column=0, sticky="w")

        
    def open_url(self, event = None):
        messagebox.showwarning("提示", "你好")


    def vip_active(self):
        pass

    def open_login(self):
        start = threading.Thread(target=login.main)
        start.start()


    def choose_explore_path(self):
        path = self.choose_directory()
        self.explore_path.delete(0, tk.END)
        self.explore_path.insert(0, path)


    def choose_explore_driver_path(self):
        path = self.choose_directory()
        self.explore_driver_path.delete(0, tk.END)
        self.explore_driver_path.insert(0, path)


    def save_explore_driver_path(self):
        self.save_explore_driver_path_config()

    def save_explore_driver_path_config(self):
        self.config['Paths_driver'] = {
            'explore_driver_path': self.explore_driver_path.get()
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def save_explore_path(self):
        self.save_explore_path_config()


    def save_explore_path_config(self):
        self.config['Paths_explore'] = {
            'explore_path': self.explore_path.get()
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def save_lhaccount(self):
        self.save_lhaccount_config()

    def save_lhaccount_config(self):
        self.config['Account'] = {
            'account': self.lhaccount.get()
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def save_lhpassword(self):
        self.save_lhpassword_config()

    def save_lhpassword_config(self):
        self.config['Password'] = {
            'password': self.lhpassword.get()
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def choose_mem_info_path(self):
        contain_mem_info = filedialog.askdirectory()
        self.config['Paths_mem_info'] = {
            'mem_info_path': contain_mem_info
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)




    def choose_org_info_path(self):
        contain_org_info = filedialog.askdirectory()
        self.config['Paths_org_info'] = {
            'org_info_path': contain_org_info 
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def choose_dev_info_path(self):
        contain_dev_info = filedialog.askdirectory()
        self.config['Paths_dev_info'] = {
            'dev_info_path': contain_dev_info
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)


    def choose_directory(self):
        path = filedialog.askopenfilename()
        return path

    def run_script(self, script_name):
        python_executable = os.path.join(os.getcwd(), 'Scripts', 'python')
        script_path = os.path.join(os.getcwd(), script_name)
        subprocess.Popen([python_executable, script_path])

    def open_mem_thread(self):
        self.config['role_mem_name'] = {
        'name_role_mem': self.role_mem.get()
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)
        temp_mem_role_name = self.config.get('role_mem_name', 'name_role_mem', fallback='')#
        if self.mem_thread is None or not self.mem_thread.is_alive():
            
            if temp_mem_role_name == "":
                messagebox.showwarning("提示", "同步数据前请先输入角色名称！要求一字不差！")
            else:
                self.mem_thread = threading.Thread(target=member.main)
                self.mem_thread.start()
                print("启动党员信息同步子程序成功！")
        else:
            print("党员信息同步子程序正在运行。")


    def close_thread_mem(self):
        member_func.stop_member_thread()
        self.mem_thread.join()
        




    def open_org_thread(self):
        self.config['role_org_name'] = {
        'name_role_org': self.role_mem.get()
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)
        temp_org_role_name = self.config.get('role_org_name', 'name_role_org', fallback='')#
        if self.org_thread is None or not self.org_thread.is_alive():
            if temp_org_role_name == "":
                messagebox.showwarning("提示", "同步数据前请先输入角色名称！要求一字不差！")
            else:
                self.org_thread = threading.Thread(target=org.main)
                self.org_thread.start()
                print("启动党组织信息同步子程序成功！")
        else:
            print("党组织信息同步子程序正在运行。")

        
    def close_thread_org(self):
        org_func.stop_org_thread()
        self.org_thread.join()
        
        

    def open_dev_thread(self):
        self.config['role_dev_name'] = {
        'name_role_dev': self.role_dev.get()
        }
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)
        temp_dev_role_name = self.config.get('role_dev_name', 'name_role_dev', fallback='')#
        if self.dev_thread is None or not self.dev_thread.is_alive():
            if temp_dev_role_name == "":
                messagebox.showwarning("提示", "同步数据前请先输入角色名称！要求一字不差！")
            else:
                self.dev_thread = threading.Thread(target=dev.main)
                self.dev_thread.start()
                print("启动党员发展信息同步子程序成功！")
        else:
            print("党员发展信息同步子程序正在运行。")


    def close_thread_dev(self):
        dev_func.stop_dev_thread()
        self.dev_thread.join()
        
    def on_closing(self):
        if self.mem_thread is not None and self.mem_thread.is_alive():
            self.close_thread_mem()
        if self.org_thread is not None and self.org_thread.is_alive():
            self.close_thread_org()
        if self.dev_thread is not None and self.dev_thread.is_alive():
            self.close_thread_dev()
        self.root2.destroy()



vip_info = vip_struct.VIP


def main(arg): 
    import tkinter as tk
    global vip_info
    vip_info = arg
    root2 = tk.Tk()
    app = MainApp(root2)
    root2.protocol("WM_DELETE_WINDOW", app.on_closing)
    root2.mainloop()


if __name__ == "__main__":
    main()



