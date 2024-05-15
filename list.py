import time
import threading

class Node:
    def __init__(self, data):
        self.data = data
        self.next = None
        self.privious = None


class LinkedList:
    def __init__(self, data):
        head_init = Node(data)
        self.head = head_init

    def insert(self, data):
        new_node = Node(data)
        self.head.privious = new_node
        new_node.next = self.head
        self.head = new_node
        thread_delay_del = threading.Thread(target=self.schedule_deletion,  args = (new_node,))
        thread_delay_del.start()

    def find_validate_by_phone_number(self, str):
        current = self.head
        while current:
            if current.data['phone_number'] == str:
                return current.data['validate_code']
            current = current.next
        return None

    def schedule_deletion(self, node):
        time.sleep(300)  # 5分钟后执行删除
        if node.next == None and node.privious == None:
            print("空节点不能被删除")
        elif node.privious == None and node.next is not None:
            self.head = node.next
            #print("成功删除节点，情况2")
        elif node.next == None and node.privious is not None:
            print("原始根节点不能被删除")
        elif node.next is not None and node.privious is not None:
            node.privious.next = node.next
            node.next.privious = node.privious
           # print("成功删除节点：情况4")

    def traverse(self):
        current = self.head
        while current:
            print(current.data)
            current = current.next

            





