from typing import List
class TreeNode:
    def __init__(self, value = None , index = 1):
        self.value = value
        self.children = []
        self.parent = []  # 父节点引用
        self.index = index

    def add_child(self, child_node):
        child_node.parent = self  # 设置父节点引用
        self.children.append(child_node)

    def dayin(self, level=0):
        ret = "\t" * level + self.value + f"{self.index}" + "\n" 
        print(ret)
        for child in self.children:
            new_level = level + 1
            
            if child.children != []:
                TreeNode.dayin(self=child, level=new_level)
            else:
                ret = "\t" * new_level + child.value + f"{child.index}"+ "\n"
                print(ret)
        

    def find_node(self, value):
        if self.value == value:
            return self
        for child in self.children:
            result = child.find_node(value)
            if result:
                return result
        return None

    def find_node_by_index(self, value):
        if self.index == value:
            return self
        for child in self.children:
            result = child.find_node_by_index(value)
            if result:
                return result
        return None

    def path_to_root(self):
        path = []
        current = self
        while current:
            path.append(current.value)
            current = current.parent
        return path[::-1]
    
    def set_value(self, value):
        self.value = value

    def copy_to_list(self, list:List):
        for child in self.children:
            
            if child.children != []:
                TreeNode.copy_to_list(self=child, list=list)
            else:
                ret = child.value
                list.append(ret)