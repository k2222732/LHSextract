import socket
import json
import threading
import mysql.connector
import datetime
import send_sms
import traceback
import list
lock = threading.Lock()
validate_sent = list.LinkedList("")

def handle_send_validate(client_socket, data):
    try:
        arg = [data['phone_number']]
        x = send_sms.Sample
        code = x._main(arg)
        code_prepare_to_send = {'validate_code':code}
        client_socket.sendall(json.dumps(code_prepare_to_send).encode('utf-8'))
        data = {'phone_number':arg[0], 'validate_code':code}
        print(data)
        global validate_sent
        validate_sent.insert(data)


    except Exception as e:
        timenow = datetime.datetime.now()
        print("handle_send_validate()发送验证码时产生异常,错误发生在", timenow, "\n", e)
        traceback.print_exc()
        client_socket.close()
        


# 连接到 MySQL 数据库（请替换为你的实际数据库信息）
                 #访客套接字     #解码后的数据
def handle_login(client_socket, data):
    try:
        username = data['username']
        password = data['password']
    except:
        timenow = datetime.datetime.now()
        print("handle_login()读取username、password键值时产生异常,错误发生在", timenow)
        traceback.print_exc()
        client_socket.close()

    try:
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
                party_organization VARCHAR(255),
                vip VARCHAR(10),
                vip_start_time DATE,
                vip_type VARCHAR(20),
                vip_deadline DATE
                )
            ''')
        db.commit()
        cursor.execute('SELECT * FROM users WHERE username=%s AND password=%s', (username, password))
        user = cursor.fetchone()
    except:
        traceback.print_exc()
        client_socket.close()
        db.close()
    
        
        
    if user:
        cursor.execute("SELECT vip,vip_start_time,vip_type,vip_deadline FROM users WHERE username = %s", (username,))
        result = cursor.fetchone()
        vip = result[0]
        vip_start_time = result[1]
        vip_type = result[2]
        vip_deadline = result[3]
        response = {'status': 'success', 'message': '登录成功','vip': vip, 'vip_start_time':vip_start_time, 'vip_type':vip_type, 'vip_deadline':vip_deadline }
    else:
        response = {'status': 'failure', 'message': '无效的用户名或密码'}
    try:
        client_socket.sendall(json.dumps(response).encode('utf-8'))
    except:
        timenow = datetime.datetime.now()
        print("handle_login()回发信息时产生异常,错误发生在", timenow)
        traceback.print_exc()
        client_socket.close()
    db.close()

def handle_register(client_socket, data):
    try:
        username = data['username']
        password = data['password']
    except:
        timenow = datetime.datetime.now()
        print("handle_register()读取键值1组时产生异常,错误发生在", timenow)
        traceback.print_exc()
        client_socket.close()

    try:
        
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
                party_organization VARCHAR(255),
                vip VARCHAR(10),
                vip_start_time DATE,   
                vip_type VARCHAR(20),
                vip_deadline DATE
                )
            ''')
        db.commit()
        cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
        existing_user = cursor.fetchone()
    except:
        traceback.print_exc()
        client_socket.close()
        db.close()
    

    if existing_user:
        response = {'status': 'failure', 'message': '用户已存在!'}
    else:
        try:
            confirm_password = data['confirm_password']
            party_organization = data.get('party_organization', '')
            phone_number = data.get('phone_number', '')
            validate_code = data.get('validate_code', '')
        except:
            timenow = datetime.datetime.now()
            print("handle_register()读取键值2组时产生异常,错误发生在", timenow)
            traceback.print_exc()
            client_socket.close()

        
        validate_code_find = validate_sent.find_validate_by_phone_number(phone_number)

        if password == confirm_password and validate_code == validate_code_find:
            cursor.execute('INSERT INTO users (username, password, phone_number, party_organization, vip, vip_start_time ,vip_type, vip_deadline) VALUES (%s, %s, %s, %s, "非会员", null, "无", null)',
                           (username, password, phone_number, party_organization))
            db.commit()
            response = {'status': 'success', 'message': '注册成功！'}
        else:
            response = {'status': 'failure', 'message': '确认密码与密码不一致或验证码错误!'}
    try:
        client_socket.sendall(json.dumps(response).encode('utf-8'))
    except:
        timenow = datetime.datetime.now()
        print("handle_register()回发信息时产生异常,错误发生在", timenow)
        traceback.print_exc()
        client_socket.close()
    db.close()


def handle_client(client_socket):
    #尝试接收数据
    request = None
    try:
        data = client_socket.recv(1024)
    except:
        timenow = datetime.datetime.now()
        print("从访客套接字中接收数据时产生异常,错误发生在", timenow)  
        traceback.print_exc()      
        client_socket.close()   
    #如果是数据包为空推出循环
    #尝试用utf-8解码数据
    try:
        request = json.loads(data.decode('utf-8'))
    except:
        timenow = datetime.datetime.now()
        print("使用utf-8解码数据时产生异常,错误发生在", timenow)
        traceback.print_exc()
    #尝试判断访客请求类型，并路由处理函数。
        client_socket.close()
    try:
        if 'action' in request:
            if request['action'] == 'login':
                handle_login(client_socket, request)
            elif request['action'] == 'register':
                handle_register(client_socket, request)
            elif request['action'] == 'call_validate_code':
                handle_send_validate(client_socket, request)
        else:
            print("action不存在login和register")
    except Exception as e:
        timenow = datetime.datetime.now()
        print("路由请求时产生异常,错误发生在", timenow)
        if request is not None:
            print(f"{request}")
        print(e)
        traceback.print_exc()
        client_socket.close()


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
            traceback.print_exc()

if __name__ == "__main__":
    main()
