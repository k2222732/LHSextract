import socket
import json
import threading
import mysql.connector
import datetime
import os
import sys
import random
from typing import List
from alibabacloud_dysmsapi20170525.client import Client as Dysmsapi20170525Client
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_dysmsapi20170525 import models as dysmsapi_20170525_models
from alibabacloud_tea_util import models as util_models
from alibabacloud_tea_util.client import Client as UtilClient
lock = threading.Lock()



class Sample:
    def __init__(self):
        pass

    @staticmethod
    def create_client() -> Dysmsapi20170525Client:
        """
        使用AK&SK初始化账号Client
        @return: Client
        @throws Exception
        """
        # 工程代码泄露可能会导致 AccessKey 泄露，并威胁账号下所有资源的安全性。以下代码示例仅供参考。
        # 建议使用更安全的 STS 方式，更多鉴权访问方式请参见：https://help.aliyun.com/document_detail/378659.html。
        config = open_api_models.Config(
            # 必填，请确保代码运行环境设置了环境变量 ALIBABA_CLOUD_ACCESS_KEY_ID。,
            access_key_id=os.environ['ALIBABA_CLOUD_ACCESS_KEY_ID'],
            # 必填，请确保代码运行环境设置了环境变量 ALIBABA_CLOUD_ACCESS_KEY_SECRET。,
            access_key_secret=os.environ['ALIBABA_CLOUD_ACCESS_KEY_SECRET']
        )
        # Endpoint 请参考 https://api.aliyun.com/product/Dysmsapi
        config.endpoint = f'dysmsapi.aliyuncs.com'
        return Dysmsapi20170525Client(config)

    @staticmethod
    def _main(
        args: List[str]
    ) -> str:
        client = Sample.create_client()
        code = Sample.generate_verification_code()
        send_sms_request = dysmsapi_20170525_models.SendSmsRequest(
            sign_name='lhsserver',
            template_code='SMS_465901623',
            phone_numbers='19953722937',
            template_param='{"code":' + code + '}'
        )
        runtime = util_models.RuntimeOptions()
        try:
            # 复制代码运行请自行打印 API 的返回值
            client.send_sms_with_options(send_sms_request, runtime)
        except Exception as error:
            # 此处仅做打印展示，请谨慎对待异常处理，在工程项目中切勿直接忽略异常。
            # 错误 message
            print(error.message)
            # 诊断地址
            print(error.data.get("Recommend"))
            UtilClient.assert_as_string(error.message)
        return code

    def generate_verification_code():
        code = ""
        for _ in range(6):
            code += str(random.randint(0, 9))
        return code



# 连接到 MySQL 数据库（请替换为你的实际数据库信息）
                 #访客套接字     #解码后的数据

def handle_login(client_socket, data, db, cursor):
    try:
        username = data['username']
        password = data['password']
    except:
        timenow = datetime.datetime.now()
        print("handle_login()读取username、password键值时产生异常,错误发生在", timenow)
    cursor.execute('SELECT * FROM users WHERE username=%s AND password=%s', (username, password))
    user = cursor.fetchone()
    if user:
        response = {'status': 'success', 'message': 'Login successful!'}
    else:
        response = {'status': 'failure', 'message': 'Invalid username or password!'}
    try:
        client_socket.sendall(json.dumps(response).encode('utf-8'))
    except:
        timenow = datetime.datetime.now()
        print("handle_login()回发信息时产生异常,错误发生在", timenow)

def handle_register(client_socket, data, db, cursor):
    try:
        username = data['username']
        password = data['password']
    except:
        timenow = datetime.datetime.now()
        print("handle_register()读取键值1组时产生异常,错误发生在", timenow)
    cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
    existing_user = cursor.fetchone()

    if existing_user:
        response = {'status': 'failure', 'message': 'Username already exists!'}
    else:
        try:
            confirm_password = data['confirm_password']

            party_organization = data.get('party_organization', '')
            phone_number = data.get('phone_number', '')
        except:
            timenow = datetime.datetime.now()
            print("handle_register()读取键值2组时产生异常,错误发生在", timenow)

        if password == confirm_password:

            cursor.execute('INSERT INTO users (username, password, phone_number, party_organization) VALUES (%s, %s, %s, %s)',
                           (username, password, phone_number, party_organization))
            db.commit()
            response = {'status': 'success', 'message': 'Registration successful!'}
        else:
            response = {'status': 'failure', 'message': 'Passwords do not match!'}
    try:

        client_socket.sendall(json.dumps(response).encode('utf-8'))
    except:
        timenow = datetime.datetime.now()
        print("handle_register()回发信息时产生异常,错误发生在", timenow)

def handle_client(client_socket):
    lock.acquire()
    db = mysql.connector.connect(
    host="127.0.0.1",
    user="root",
    password="327105",
    database="LHSmanager"
    )
    cursor = db.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY,
            username VARCHAR(255) NOT NULL,
            password VARCHAR(255) NOT NULL,
            phone_number VARCHAR(15),
            party_organization VARCHAR(255)
            )
        ''')
    db.commit()
    request = None
    #尝试接收数据
    try:
        try:
            data = client_socket.recv(1024)
        except:
            timenow = datetime.datetime.now()
            print("从访客套接字中接收数据时产生异常,错误发生在", timenow)           
        #如果是数据包为空推出循环
        #尝试用utf-8解码数据
        try:
            request = json.loads(data.decode('utf-8'))
        except:
            timenow = datetime.datetime.now()
            print("使用utf-8解码数据时产生异常,错误发生在", timenow)
        #尝试判断访客请求类型，并路由处理函数。
        try:
            if 'action' in request:
                if request['action'] == 'login':
                    handle_login(client_socket, request, db, cursor)
                elif request['action'] == 'register':
                    handle_register(client_socket, request, db, cursor)
            else:
                print("action不存在login和register")
        except Exception as e:
            timenow = datetime.datetime.now()
            print("路由请求时产生异常,错误发生在", timenow)
            if request is not None:
                print(f"{request}")
            print(e)

    finally:
        lock.release()
        db.close()
    client_socket.close()
    print("退出接待线程\n******************************************")


def main():
    #创建tcp\ip协议套接字，监听8888端口，最多接收50个访问。
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_socket.bind(('', 8888))
    server_socket.listen(50)
    print("Server listening on port 8888...")
    while True:
        try:
            #等待访问，阻塞
            client_socket, addr = server_socket.accept()
            #访问来了，打印访客ip地址，启动新线程运行处理程序，此处有可能涉及多线程同步问题。
            print(f"Connection from {addr}")
            client_handler = threading.Thread(target=handle_client, args=(client_socket,))
            client_handler.start()
        except:
            timenow = datetime.datetime.now()
            print("堵塞等待套接字，启动新线程时产生异常,错误发生在", timenow)

if __name__ == "__main__":
    main()
