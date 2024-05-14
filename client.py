import tkinter as tk
from tkinter import messagebox
import json
import socket
import threading
import time


class ClientApp:
    def __init__(self, root):
        self.ui_thread = None
        self.root = root
        self.root.title("组工填表神器")
        self.label_username = tk.Label(root, text="用户名:")
        self.label_password = tk.Label(root, text="密码:")
        self.entry_username = tk.Entry(root)
        self.entry_password = tk.Entry(root, show="*")
        self.label_username.grid(row=0, sticky=tk.E, pady=15)
        self.label_password.grid(row=1, sticky=tk.E, pady=5)
        self.entry_username.grid(row=0, column=1, pady=5)
        self.entry_password.grid(row=1, column=1, pady=5)
        self.login_button = tk.Button(root, text="登录", command=self.login)
        self.register_button = tk.Button(root, text="注册", command=self.register)
        self.login_button.grid(row=2, column=1, sticky=tk.E, pady=5)
        self.register_button.grid(row=2, column=2, sticky=tk.W, pady=5)
        self.remaining_time = tk.StringVar()


    def login(self):
        username = self.entry_username.get()
        password = self.entry_password.get()
        request = {'action': 'login', 'username': username, 'password': password}
        self.send_request(request)


    def call_validate_code(self):
        if not self.entry_phonenum_r.get():
            phone_number = self.entry_phonenum_r.get()
            request = {'action':'call_validate_code', 'phone_number':phone_number}
            self.send_request(request)
            self.send_validate_code.config(state=tk.DISABLED)
            self.remaining_time.set("60")
            threading.Thread(target=self.enable_button_after_delay).start()
        else:
            messagebox.showinfo("请输入手机号码！")


    def enable_button_after_delay(self):
        for i in range(60, 0, -1):  # 倒计时60秒
            self.remaining_time.set(str(i))  # 更新剩余时间显示
            time.sleep(1)
        self.send_validate_code.config(state=tk.NORMAL)  # 恢复按钮为可用状态
        self.remaining_time.set("")


    def register(self):
        self.root.withdraw()
        self.registration_window = tk.Toplevel(self.root)
        self.registration_window.protocol("WM_DELETE_WINDOW", self.show_login_window)
        self.label_username_r = tk.Label(self.registration_window, text="用户名:")
        self.label_password_r = tk.Label(self.registration_window, text="密码:")
        self.label_cfpassword_r = tk.Label(self.registration_window, text="确认密码:")
        self.label_phonenum_r = tk.Label(self.registration_window, text="手机号码:")
        self.validate_code = tk.Entry(self.registration_window, width = 10)
        self.send_validate_code = tk.Button(self.registration_window, text="发送验证码", command= self.call_validate_code)
        self.time_left = tk.Label(self.registration_window, textvariable=self.remaining_time)
        self.label_org_r = tk.Label(self.registration_window, text="所在党支部:")
        self.entry_username_r = tk.Entry(self.registration_window)
        self.entry_password_r = tk.Entry(self.registration_window, show="*")
        self.entry_cfpassword_r = tk.Entry(self.registration_window, show="*")
        self.entry_phonenum_r = tk.Entry(self.registration_window)
        self.entry_org_r = tk.Entry(self.registration_window)
        self.label_username_r.grid(row=0, sticky=tk.E, pady=5)
        self.label_password_r.grid(row=1, sticky=tk.E, pady=5)
        self.label_cfpassword_r.grid(row=2, sticky=tk.E, pady=5)
        self.label_phonenum_r.grid(row=3, sticky=tk.E, pady=5)
        self.time_left.grid(row=3, sticky=tk.E, pady=5)

        self.validate_code.grid(row=4, column=1, sticky=tk.E, padx=(0,7))
        self.send_validate_code.grid(row=4, column=2, sticky=tk.E, pady=5)
        self.time_left.grid(row=4, column=3, sticky=tk.E, pady=5)

        self.label_org_r.grid(row=5, sticky=tk.E, pady=5)
        self.entry_username_r.grid(row=0, column=1, padx=5, pady=5)
        self.entry_password_r.grid(row=1, column=1, padx=5, pady=5)
        self.entry_cfpassword_r.grid(row=2, column=1, padx=5, pady=5)
        self.entry_phonenum_r.grid(row=3, column=1, padx=5, pady=5)

        self.entry_org_r.grid(row=5, column=1, padx=5, pady=5)
        self.register_button_r = tk.Button(self.registration_window, text="注册", command=self.reg_confirm)
        self.cancel_button_r = tk.Button(self.registration_window, text="取消", command=self.reg_cancel)
        self.register_button_r.grid(row=6, column=0, padx=0, pady=5)
        self.cancel_button_r.grid(row=6, column=1, padx=0, pady=5)


    def reg_confirm(self):
        user_name = self.entry_username_r.get()
        password = self.entry_password_r.get()
        cfpassword = self.entry_cfpassword_r.get()
        phonenum = self.entry_phonenum_r.get()
        org = self.entry_org_r.get()
        validate_code = self.validate_code.get()
        request =  {'action': 'register', 'username': user_name, 'password': password, 'confirm_password':cfpassword, 'phone_number':phonenum, 'validate_code':validate_code, 'party_organization':org}
        self.send_request(request)
        self.registration_window.destroy()
        self.root.deiconify()


    def reg_cancel(self):
        self.registration_window.destroy()
        self.root.deiconify()


    def show_login_window(self):
        self.registration_window.destroy()
        self.root.deiconify()


    def send_request(self, request):
        try:
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.connect(('82.157.124.132', 8888))
            client_socket.sendall(json.dumps(request).encode('utf-8'))
            response_data = client_socket.recv(1024)
            response = json.loads(response_data.decode('utf-8'))
            if request['action'] == 'login':

                if response['status'] == 'success':
                    messagebox.showinfo("登录成功", response['message'])
                    self.open_ui_thread()
                    self.root.destroy()
                else:
                    messagebox.showerror("登录失败", response['message'])

            elif request['action'] == 'register':

                if response['status'] == 'success':
                    messagebox.showinfo("注册成功", response['message'])
                else:
                    messagebox.showerror("注册失败", response['message'])
        finally:
            if client_socket:
                client_socket.close()


    def open_ui_thread(self):
        if self.ui_thread is None or not self.ui_thread.is_alive():
            self.ui_thread = threading.Thread(target=self.load_ui)
            self.ui_thread.start()
            print("UI界面启动成功！")
        else:
            print("UI线程正在工作。")


    def load_ui(self):
        import ui


if __name__ == "__main__":
    root = tk.Tk()
    app = ClientApp(root)
    root.mainloop()


