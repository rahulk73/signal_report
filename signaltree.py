from sqlscript import GetSignals
class Node:
    def __init__(self,data,parent=None):
        self.data=data
        self.children=[]
        if parent:
            parent.children.append(self)
    def getChildrenData(self):
        child_names=[]
        for child in self.children:
            child_names.append(child.data)
        return child_names
    def getChildren(self):
        return [child for child in self.children]
    def findChild(self,data):
        for child in self.children:
            if data==child.data:
                return child
    def isLeaf(self):
        return self.children==[]
    def isInternalNode(self):
        return not self.children==[]
    def __str__(self):
        return ', '.join([child.data for child in self.children])
 
class Tree:
    def __init__(self,data_tuples):
        def helper(next_parent,data_split,absolute_path):
            for node_val in data_split:
                if node_val not in next_parent.getChildrenData():
                    next_parent=Node(node_val,parent=next_parent)
                else:
                    next_parent=next_parent.findChild(node_val)
            next_parent.absolute_path=absolute_path

        self.root=[Node('Electrical'),Node('System'),Node('Electrical-Control')]
        for data in data_tuples:
            data_split=data[0].split(' / ')
            if data[1]=='site':
                if data[2]:
                    if data[2][-1]!='c':
                        helper(self.root[0],data_split,data[0])
                    else:
                        helper(self.root[2],data_split,data[0])
            else:
                helper(self.root[1],data_split,data[0])
    def getRoot(self):
        return self.root
if __name__ == "__main__":
    signals=GetSignals()
    tree=Tree(signals.result)
    print(tree.getRoot()[0])
