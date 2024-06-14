
class TreeNode:
    def __init__(self, value = None, parent=None):
        self.value = value
        self.children = []
        self.parent = parent  # 父节点引用

    def add_child(self, child_node):
        child_node.parent = self  # 设置父节点引用
        self.children.append(child_node)

    def __repr__(self, level=0):
        ret = "\t" * level + repr(self.value) + "\n"
        for child in self.children:
            ret += child.__repr__(level + 1)
        return ret

    def find_node(self, value):
        if self.value == value:
            return self
        for child in self.children:
            result = child.find_node(value)
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