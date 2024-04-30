import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import configparser
import os
import time


class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("灯塔填表小助手")
        self.tabControl = ttk.Notebook(root)
        self.privious_tab = ""
        self.config_file = 'config.ini'
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file)

        # 选项卡1
        self.tab1 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab1, text="基本设置")
        self.setup_tab1()
        self.load_config()

        # Binding the "q" key to delete password field
        self.root.bind('q', self.delete_password_field)

    def load_config(self):
        if os.path.exists(self.config_file):
            self.explore_driver_path.insert(0, self.config.get('Paths_driver', 'explore_driver_path', fallback=''))
            self.explore_path.insert(0, self.config.get('Paths_explore', 'explore_path', fallback=''))
            self.lhaccount.insert(0, self.config.get('Account', 'account', fallback=''))
            self.lhpassword.insert(0, self.config.get('Password', 'password', fallback=''))
        else:
            pass

    def setup_tab1(self):
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

        self.tabControl.grid(row=0, column=0, sticky="nsew")

    def delete_password_field(self, event=None):
        # Deleting password label, entry, and save button
        n = 1
        num_rows, num_columns = self.tab1.grid_size()
        while n <= num_rows-1:
            slaves = self.tab1.grid_slaves(row=n)
            for slave in slaves:
                slave.destroy()
            n = n + 1
        self.setup_tab1()

    def choose_explore_driver_path(self):
        path = self.choose_directory()
        self.explore_driver_path.delete(0, tk.END)
        self.explore_driver_path.insert(0, path)

    def save_explore_driver_path(self):
        self.save_explore_driver_path_config()

    def choose_explore_path(self):
        path = self.choose_directory()
        self.explore_path.delete(0, tk.END)
        self.explore_path.insert(0, path)

    def choose_directory(self):
        path = filedialog.askopenfilename()
        return path

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


root = tk.Tk()
app = MainApp(root)
root.mainloop()
