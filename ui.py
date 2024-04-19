import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import os
import subprocess
import configparser
import threading

role_name_mem = ""
role_name_org = ""
role_name_dev = ""

##############
class MainApp:
    def __init__(self, root):
        self.mem_thread = None
        self.org_thread = None
        self.dev_thread = None
        self.config_file = 'config.ini'
        self.config = configparser.ConfigParser()


        self.root = root
        self.root.title("灯塔填表小助手")
        self.tabControl = ttk.Notebook(root)


        # 选项卡1
        self.tab1 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab1, text="基本设置")
        self.setup_tab1()

        # 选项卡2
        self.tab2 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab2, text="仅登录")
        self.setup_tab2()


        # 选项卡3
        self.tab3 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab3, text="数据同步")
        self.setup_tab3()

        # 选项卡4
        self.tab4 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab4, text="主题党日上传")
        self.setup_tab4()

        # 选项卡5
        self.tab5 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab5, text="发展党员材料上传")
        self.setup_tab5()


        # 选项卡6
        self.tab6 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab6, text="自动填表")
        self.setup_tab6()


        # 选项卡7
        self.tab6 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab6, text="账号信息")
        self.setup_tab7()


        self.tabControl.pack(expand=1, fill="both")
        self.load_config()
    

    def load_config(self):
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
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
        pass
        


    def setup_tab3(self):
        ttk.Button(self.tab3, text="同步党员信息", command=self.open_mem_thread).grid(row=0, column=1, pady=5, padx=(10, 5))
        ttk.Button(self.tab3, text="选择目录", command=self.choose_mem_info_path).grid(row=0, column=2, pady=5, padx=(10, 5))
        self.role_mem = ttk.Entry(self.tab3, state="normal")
        self.role_mem.grid(row=0, column=3, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        self.role_mem.insert(0, "请输入需下载党员信息的角色(要求名称严格一致)")
        ttk.Button(self.tab3, text="同步党组织信息", command=self.open_org_thread).grid(row=1, column=1, pady=5, padx=(10, 5))
        ttk.Button(self.tab3, text="选择目录", command=self.choose_org_info_path).grid(row=1, column=2, pady=5, padx=(10, 5))
        self.role_org = ttk.Entry(self.tab3, state="normal")
        self.role_org.grid(row=1, column=3, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        self.role_org.insert(0, "请输入需下载党组织信息的角色(要求名称严格一致)")
        ttk.Button(self.tab3, text="同步发展纪实信息", command=self.open_dev_thread).grid(row=2, column=1, pady=5, padx=(10, 5))
        ttk.Button(self.tab3, text="选择目录", command=self.choose_dev_info_path).grid(row=2, column=2, pady=5, padx=(10, 5))
        self.role_dev = ttk.Entry(self.tab3, state="normal")
        self.role_dev.grid(row=2, column=3, columnspan=2, sticky="ew", pady=5, padx=(10, 0))
        self.role_dev.insert(0, "请输入需下载党员发展信息的角色(要求名称严格一致)")
        


    def setup_tab4(self):
        pass


    def setup_tab5(self):
        pass


    def setup_tab6(self):
        pass

    def setup_tab7(self):
        pass


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
        if self.mem_thread is None or not self.mem_thread.is_alive():
            if self.role_mem.get() == "请输入需下载党员信息的角色(要求名称严格一致)" or None:
                messagebox.showwarning("提示", "同步数据前请先输入角色名称！")
            else:
                global role_name_mem
                role_name_mem = self.role_mem.get()
                self.mem_thread = threading.Thread(target=self.open_mem)
                self.mem_thread.start()
                print("启动党员信息同步子程序成功！")
        else:
            print("党员信息同步子程序正在运行。")

    def open_mem(self):
        import member



    def open_org_thread(self):
        if self.org_thread is None or not self.org_thread.is_alive():
            if self.role_org.get() == "请输入需下载党组织信息的角色(要求名称严格一致)" or None:
                messagebox.showwarning("提示", "同步数据前请先输入角色名称！")
            else:
                global role_name_org
                role_name_org = self.role_mem.get()
                self.org_thread = threading.Thread(target=self.open_org)
                self.org_thread.start()
                print("启动党组织信息同步子程序成功！")
        else:
            print("党组织信息同步子程序正在运行。")

        
    def open_org(self):
        import org
    


    def open_dev_thread(self):
        if self.dev_thread is None or not self.dev_thread.is_alive():
            if self.role_dev.get() == "请输入需下载党员发展信息的角色(要求名称严格一致)" or None:
                messagebox.showwarning("提示", "同步数据前请先输入角色名称！")
            else:
                global role_name_dev
                role_name_dev = self.role_dev.get()
                self.dev_thread = threading.Thread(target=self.open_dev)
                self.dev_thread.start()
                print("启动党员发展信息同步子程序成功！")
        else:
            print("党员发展信息同步子程序正在运行。")

        
    def open_dev(self):
        import dev

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
