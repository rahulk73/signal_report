from sqlscript import GetSignals
class Node:
    def __init__(self,data,parent=None):
        self.data=data
        self.children=[]
        self.absolute_path=''
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

        self.control = frozenset(['modulespc','switch_dpc','cs_ctrlonoff_spc','moduledpc'])
        self.allsignals = frozenset(['mappingsps','cs_voltageabsence_sps','cs_authostate_sps','cs_onoff_sps','cs_voltagerefpresence_sp','computedswitchpos_dps',
        'cs_busbarvchoice_sps','cs_acceptforcing_sps','groupsps','tapfct','cs_voltagepresence_sps','cs_closeorderstate_sps','userfunctionsps',
        'cs_voltagerefabsence_sps','moduledps','modulesps'])
        self.measurement = frozenset(['modulemv'])
        self.meter = frozenset(['modulecounter'])
        self.root={'All':{'site':Node('Electrical'),'scs':Node('System')}, 'Control':Node('Electrical-Control'), 'Measurement':Node('Electrical-Measurement'),
        'Meter':Node('Electrical-Meter')}
        for fullpath,typ0,typ5,iec in data_tuples:
            data_split=fullpath.split(' / ')
            if ((not typ5 and iec) or (typ5 in self.allsignals)) and typ0:
                helper(self.root['All'][typ0],data_split,fullpath)
            elif typ5 in self.control:
                helper(self.root['Control'],data_split,fullpath)
            elif typ5 in self.measurement:
                helper(self.root['Measurement'],data_split,fullpath)
            elif typ5 in self.meter:
                helper(self.root['Meter'],data_split,fullpath)
    def getRoot(self):
        return self.root
if __name__ == "__main__":
    signals=GetSignals()
    tree=Tree(signals.result)
    print(tree.getRoot()['Meter'].isInternalNode())
