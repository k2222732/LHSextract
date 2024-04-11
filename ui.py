import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import os
import subprocess
import configparser
import globalv

class MainApp:
    def __init__(self, root):
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
        ttk.Button(self.tab3, text="同步党员信息", command=self.sync_member_info).pack(pady=10)
        ttk.Button(self.tab3, text="同步党组织信息", command=self.sync_organization_info).pack(pady=10)
        ttk.Button(self.tab3, text="同步发展纪实信息", command=self.sync_development_info).pack(pady=10)
        


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


    def sync_member_info(self):
        self.run_script("member.py")

    def sync_organization_info(self):
        self.run_script("org.py")

    def sync_development_info(self):
        self.run_script("dev.py")

    def choose_directory(self):
        path = tkinter.filedialog.askopenfilename()
        return path

    def run_script(self, script_name):
        script_path = os.path.join(os.getcwd(), script_name)
        subprocess.Popen(['python', script_path])

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
