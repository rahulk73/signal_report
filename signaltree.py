import pymysql.cursors
def get_data():
    result=()
    connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
    try:
        with connection.cursor() as cursor:
            sql="SELECT object_fullpath FROM objects"
            cursor.execute(sql)
            cursor.fetchmany()
            result=cursor.fetchall()
    except Exception as e:
        print(e)
        return -2
    finally:
        connection.close()
        return result
class Node:
    def __init__(self,data,parent=None):
        self.data=data
        self.children=[]
        if parent:
            parent.children.append(self)
    def getChildren(self):
        child_names=[]
        for child in self.children:
            child_names.append(child.data)
        return child_names
    def findChild(self,data):
        for child in self.children:
            if data==child.data:
                return child
    def isLeaf(self):
        return self.children==[]
    def __str__(self):
        return ', '.join([child.data for child in self.children])
 
class Tree:
    def __init__(self,data_tuples):
        self.root=Node('root')
        for data in data_tuples:
            next_parent=self.root
            data_split=data[0].split(' / ')
            for node_val in data_split:
                if node_val not in next_parent.getChildren():
                    next_parent=Node(node_val,parent=next_parent)
                else:
                    next_parent=next_parent.findChild(node_val)
    def getRoot(self):
        return self.root
if __name__ == "__main__":
    tree=Tree(get_data())
    print(tree.getRoot())
