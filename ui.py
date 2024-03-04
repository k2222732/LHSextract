import tkinter as tk
from tkinter import ttk
import os
import subprocess

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("灯塔填表小助手")

        self.tabControl = ttk.Notebook(root)

        # 选项卡1
        self.tab1 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab1, text="基本设置")

        self.setup_tab1()

        # 选项卡2
        self.tab2 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab2, text="数据同步")

        self.setup_tab2()

        # 选项卡3
        self.tab3 = ttk.Frame(self.tabControl)
        self.tabControl.add(self.tab3, text="选项卡3")

        self.tabControl.pack(expand=1, fill="both")

    def setup_tab1(self):
        # 创建 Label 和 Entry 用于显示和选择目录
        ttk.Label(self.tab1, text="党员信息目录：").grid(row=0, column=0, sticky="e", pady=5)
        self.member_info_path_entry = ttk.Entry(self.tab1, state="readonly")
        self.member_info_path_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Button(self.tab1, text="选择目录", command=self.choose_member_info_path).grid(row=0, column=3, pady=5)

        ttk.Label(self.tab1, text="党组织信息目录：").grid(row=1, column=0, sticky="e", pady=5)
        self.organization_info_path_entry = ttk.Entry(self.tab1, state="readonly")
        self.organization_info_path_entry.grid(row=1, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Button(self.tab1, text="选择目录", command=self.choose_organization_info_path).grid(row=1, column=3, pady=5)

        ttk.Label(self.tab1, text="党员发展纪实目录：").grid(row=2, column=0, sticky="e", pady=5)
        self.development_info_path_entry = ttk.Entry(self.tab1, state="readonly")
        self.development_info_path_entry.grid(row=2, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Button(self.tab1, text="选择目录", command=self.choose_development_info_path).grid(row=2, column=3, pady=5)

    def setup_tab2(self):
        ttk.Button(self.tab2, text="同步党员信息", command=self.sync_member_info).pack(pady=10)
        ttk.Button(self.tab2, text="同步党组织信息", command=self.sync_organization_info).pack(pady=10)
        ttk.Button(self.tab2, text="同步发展纪实信息", command=self.sync_development_info).pack(pady=10)

    def choose_member_info_path(self):
        path = self.choose_directory()
        self.member_info_path_entry.delete(0, tk.END)
        self.member_info_path_entry.insert(0, path)

    def choose_organization_info_path(self):
        path = self.choose_directory()
        self.organization_info_path_entry.delete(0, tk.END)
        self.organization_info_path_entry.insert(0, path)

    def choose_development_info_path(self):
        path = self.choose_directory()
        self.development_info_path_entry.delete(0, tk.END)
        self.development_info_path_entry.insert(0, path)

    def sync_member_info(self):
        self.run_script("member.py")

    def sync_organization_info(self):
        self.run_script("org.py")

    def sync_development_info(self):
        self.run_script("dev.py")

    def choose_directory(self):
        path = tk.filedialog.askdirectory()
        return path

    def run_script(self, script_name):
        script_path = os.path.join(os.getcwd(), script_name)
        subprocess.Popen(['python', script_path])

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
