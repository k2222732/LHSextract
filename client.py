import tkinter as tk
from tkinter import messagebox
import subprocess
import json
import socket
import time

class ClientApp:
    def __init__(self, root):
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

    def login(self):
        username = self.entry_username.get()
        password = self.entry_password.get()

        request = {'action': 'login', 'username': username, 'password': password}
        self.send_request(request)

    def register(self):
        self.root.withdraw()
        self.registration_window = tk.Toplevel(self.root)
        self.registration_window.protocol("WM_DELETE_WINDOW", self.show_login_window)

        self.label_username_r = tk.Label(self.registration_window, text="用户名:")
        self.label_password_r = tk.Label(self.registration_window, text="密码:")
        self.label_cfpassword_r = tk.Label(self.registration_window, text="确认密码:")
        self.label_phonenum_r = tk.Label(self.registration_window, text="手机号码:")
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
        self.label_org_r.grid(row=4, sticky=tk.E, pady=5)
        self.entry_username_r.grid(row=0, column=1, padx=5, pady=5)
        self.entry_password_r.grid(row=1, column=1, padx=5, pady=5)
        self.entry_cfpassword_r.grid(row=2, column=1, padx=5, pady=5)
        self.entry_phonenum_r.grid(row=3, column=1, padx=5, pady=5)
        self.entry_org_r.grid(row=4, column=1, padx=5, pady=5)

        self.register_button_r = tk.Button(self.registration_window, text="注册", command=self.reg_confirm)
        self.cancel_button_r = tk.Button(self.registration_window, text="取消", command=self.reg_cancel)
        self.register_button_r.grid(row=5, column=1, padx=5, pady=5)
        self.cancel_button_r.grid(row=5, column=2, padx=5, pady=5)
        pass

    def reg_confirm(self):
        user_name = self.entry_username_r.get()
        password = self.entry_password_r.get()
        cfpassword = self.entry_cfpassword_r.get()
        phonenum = self.entry_phonenum_r.get()
        org = self.entry_org_r.get()
        request =  {'action': 'register', 'username': user_name, 'password': password, 'confirm_password':cfpassword, 'phone_number':phonenum, 'party_organization':org}
        self.send_request(request)
        self.registration_window.destroy()
        self.root.deiconify()

    def reg_cancel(self):
        self.registration_window.destroy()
        self.root.deiconify()

    def show_login_window(self):
        self.root.deiconify()

    def send_request(self, request):
        client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        client_socket.connect(('82.157.124.132', 8888))
        client_socket.sendall(json.dumps(request).encode('utf-8'))
        response_data = client_socket.recv(1024)
        response = json.loads(response_data.decode('utf-8'))
        if request['action'] == 'login':

            if response['status'] == 'success':
                messagebox.showinfo("登录成功", response['message'])
                self.open_ui()
                self.root.destroy()
            else:
                messagebox.showerror("登录失败", response['message'])

        elif request['action'] == 'register':

            if response['status'] == 'success':
                messagebox.showinfo("注册成功", response['message'])
            else:
                messagebox.showerror("注册失败", response['message'])


        client_socket.close()


    def open_ui(self):
        subprocess.Popen(['python', 'ui.py'])

if __name__ == "__main__":
    root = tk.Tk()
    app = ClientApp(root)
    root.mainloop()
