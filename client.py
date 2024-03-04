import tkinter as tk
from tkinter import messagebox
import subprocess
import json
import socket

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
        # Your registration logic here
        pass

    def send_request(self, request):
        client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        client_socket.connect(('localhost', 8888))

        client_socket.sendall(json.dumps(request).encode('utf-8'))
        response_data = client_socket.recv(1024)
        response = json.loads(response_data.decode('utf-8'))

        if response['status'] == 'success':
            messagebox.showinfo("登录成功", response['message'])
            self.open_ui()
        else:
            messagebox.showerror("登录失败", response['message'])

        client_socket.close()

    def open_ui(self):
        subprocess.Popen(['python', 'ui.py'])

if __name__ == "__main__":
    root = tk.Tk()
    app = ClientApp(root)
    root.mainloop()
