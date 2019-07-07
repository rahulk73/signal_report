"""
Application to extract data from a table on a mysql database based on a predifned frequency interval
and load this data on an excel spreadsheet.
Credits:

Icon made by VisualPharm at icon-icons.com/icon/database-the-application/2803 (License : CC Attribution)
"""
from tkcalendar import *
from xlsscript import ParseData
from sqlscript import GetSignals
from signaltree import Tree,Node
import threading
from os import system
w=1200
h=700

class Navbar(tk.Frame):
    def __init__(self,parent):
        tk.Frame.__init__(self,parent)
        self.parent=parent
        self.internal_nodes = dict()
        self.tree = ttk.Treeview(self)
        ysb = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        xsb = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        self.tree.heading('#0', text='Signal tree', anchor='w')
        self.tree.grid(ipadx=100,ipady=100,sticky='e')
        ysb.grid(row=0, column=1, sticky='ns')
        xsb.grid(row=1, column=0, sticky='ew')
        self.OPTIONS=['All Signals','All Controls']
        self.optionvar=tk.StringVar(self)
        self.optionvar.set(self.OPTIONS[0])
        self.optionmenu=tk.OptionMenu(self,self.optionvar,*self.OPTIONS,command=self.callback)
        self.optionmenu.grid(row=0,column=2)
        self.layout=Tree(GetSignals().result)
        self.electrical=self.layout.root[0]
        self.system=self.layout.root[1]
        self.electrical_control=self.layout.root[2]
        self.root_iid=[]
        self.insert_node('', self.electrical,self.electrical.data,root=True)
        self.insert_node('', self.system,self.system.data,root=True)
        self.tree.bind('<<TreeviewOpen>>', self.open_node)
    def insert_node(self, parent_iid, node,text,root=False):
        node_iid = self.tree.insert(parent_iid, 'end', text=text, open=False)
        if node.isInternalNode():
            self.internal_nodes[node_iid] = node
            self.tree.insert(node_iid, 'end')
        else:
            self.tree.item(node_iid,values=[node.absolute_path])
        if root:
            self.root_iid.append(node_iid)
    def open_node(self, event):
        node_iid = self.tree.focus()
        node = self.internal_nodes.pop(node_iid, None)
        if node:
            self.tree.delete(self.tree.get_children(node_iid))
            for node_child in node.getChildren():
                self.insert_node(node_iid,node_child, node_child.data)
        else:
            values=self.tree.item(node_iid,option='values')
            if values:
                self.parent.plabel.config(text=values[0])
    def callback(self,*args):
        self.tree.delete(*self.root_iid)
        self.root_iid=[]
        if self.optionvar.get() == 'All Controls':
            self.insert_node('', self.electrical_control,self.electrical_control.data,root=True)
        else:
            self.insert_node('', self.electrical,self.electrical.data,root=True)
            self.insert_node('', self.system,self.system.data,root=True)
class MainApplication(tk.Canvas):
    def __init__(self,parent):
        tk.Canvas.__init__(self,parent)
        self.parent=parent
        self.setup()
        self.navbar=Navbar(self)
        self.navbar.place(x=550,y=20)
        self.create_image(0,0,image=self.photoimage,anchor='nw')
        self.create_text(15,20,text="Time Period :",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.fbutton.place(x=150,y=20)
        self.tbutton.place(x=250,y=20)
        self.flabel.place(x=150,y=45)
        self.tlabel.place(x=250,y=45)
        self.advanbutton.place(x=350,y=25)
        self.create_text(15,90,text="Object Path: ",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.plabel.place(x=150,y=90,width=370)
        self.cbutton.place(x=150,y=200)
    def extract(self):
        def thread_extract():
            self.progress.place(x=200,y=300)
            self.progress.start()
            self.object_fullpath=self.plabel.cget('text')
            if self.hidden:
                found=ParseData(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day),
                datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day),
                0,0,self.hidden,self.object_fullpath.upper(),self.navbar.optionvar.get()).result
            else:
                found=ParseData(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day),
                datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day),
                self.fqentry.get(),self.radb.get(),self.hidden,self.object_fullpath.upper(),self.navbar.optionvar.get()).result       
            self.progress.stop()
            self.progress.place_forget()
            self.enable_all()
            if found==-1:
                tk.messagebox.showerror("Error","Close the excel file and try again.")
            elif found==-2:
                tk.messagebox.showerror("Error","Something went wrong. Restart the application and try again.")
            elif found==-3:
                tk.messagebox.showwarning("Info","No data was found for the signal.")
            elif found==-4:
                tk.messagebox.showerror("Error","Too much data to process. Narrow down your search criteria and try again")
            elif found==0:
                tk.messagebox.showwarning("Info","No data was found for the signal in the selected time period.")
            elif found==1:
                tk.messagebox.showinfo("Extraction Successful !","You can view the records in the file tables.xlsx")
                system('start EXCEL.EXE ./SignalLog/tables.xlsx')
        self.disable_all()        
        threading.Thread(target=thread_extract).start()
        
    def setup(self):
        self.tdate=calendar.datetime.date.today()
        self.fdate=calendar.datetime.date(self.tdate.year,self.tdate.month,self.tdate.day)+datetime.timedelta(days=-2)
        self.fbutton=ttk.Button(self, text='From', command=self.fgetdate,underline=1)
        self.tbutton=ttk.Button(self, text='To', command=self.tgetdate)
        self.tbutton.config(state='disabled')
        self.fstr=tk.StringVar(self,self.date_format(self.fdate))
        self.flabel=tk.Label(self,textvariable=self.fstr,width=10)
        self.tstr=tk.StringVar(self,self.date_format(self.tdate))
        self.tlabel=tk.Label(self,textvariable=self.tstr,width=10)
        self.photoimage=tk.PhotoImage(file="C:\\Users\\OISM\\Desktop\\sqlApp\\bgimage.png")
        self.parent.geometry("%dx%d" % (w,h))
        self.parent.title("Create Exel Log File")
        self.btn_text=tk.StringVar(self,value="Show Advanced Options")
        self.advanbutton=tk.Button(self,textvariable=self.btn_text,command=self.advance)
        self.hidden=True
        self.fqvar=tk.StringVar(self,value='1')
        self.fqvar.trace('w',self.update)
        self.radb=tk.IntVar(self,2)
        self.flag=tk.IntVar(self,0)
        self.flag.trace('w',self.callback)
        self.default_text='MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN'
        self.plabel=tk.Label(self,text=self.default_text)
        self.cbutton=tk.Button(self,text="Create excel File!",command=self.extract,width=20,height=2)
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL,length=100,  mode='indeterminate')
    def fgetdate(self):
        def print_sel():
            if self.cal.selection_get()<calendar.datetime.date.today():
                self.flag.set(1)
                self.fdate=self.cal.selection_get()
                self.fstr.set(self.date_format(self.fdate))
                self.top.destroy()
            else:
                messagebox.showerror("Date Error","Date is invalid. Try Again")

        self.top = tk.Toplevel(self)
        self.top.grab_set()

        self.cal = Calendar(self.top, font="Arial 14", selectmode='day',
                        cursor="hand1",year=self.fdate.year,month=self.fdate.month,day=self.fdate.day)

        self.cal.pack(fill="both", expand=True)
        ttk.Button(self.top, text="Go", command=print_sel).pack()
    def tgetdate(self):
        def print_sel():
            if self.cal.selection_get()>self.fdate and self.cal.selection_get()<=calendar.datetime.date.today():
                self.flag.set(0)
                self.tdate=self.cal.selection_get()
                self.tstr.set(self.date_format(self.tdate))
                self.top.destroy()
            else:
                messagebox.showerror("Date Error","The time interval is invalid. Try Again")
        self.top = tk.Toplevel(self)
        self.top.grab_set()
        self.cal = Calendar(self.top, font="Arial 14", selectmode='day',
                        cursor="hand1",year=self.tdate.year,month=self.tdate.month,day=self.tdate.day)
        self.cal.pack(fill="both", expand=True)
        ttk.Button(self.top, text="Go", command=print_sel).pack()
    def advance(self):
        if self.hidden:
            self.radiobuttton1=tk.Radiobutton(self,text="Second(s)",value=1,variable=self.radb)
            self.radiobuttton2=tk.Radiobutton(self,text="Minute(s)",value=2,variable=self.radb)
            self.radiobuttton3=tk.Radiobutton(self,text="Hour(s)",value=3,variable=self.radb)
            self.fqentry=tk.Entry(self,width=2,textvariable=self.fqvar)
            self.reg=self.register(self.valid)
            self.fqentry.config(validate='key',validatecommand=(self.reg,'%P'))
            self.btn_text.set("Hide Advanced Options")
            self.create_text(15,130,text="Show changes every ",fill='white',anchor='nw',font=("Arial",12,'bold'),tag='showtext')  
            self.fqentry.place(x=180,y=133)
            self.radiobuttton1.place(x=200,y=130)
            self.radiobuttton2.place(x=280,y=130)
            self.radiobuttton3.place(x=360,y=130)
        else:
            self.btn_text.set("Show Advanced Options")
            self.delete('showtext')
            self.delete('defaultinfo')
            self.fqentry.place_forget()
            self.radiobuttton1.place_forget()
            self.radiobuttton2.place_forget()
            self.radiobuttton3.place_forget()
        self.hidden = not self.hidden

    def callback(self,*args):
        if(self.flag.get()):
            print("To button enabled",args)
            self.tbutton.config(state='normal')
    def update(self,*args):
        if(self.fqvar.get()):
            self.frequency=int(self.fqvar.get())
        else:
            self.frequency=1
    def valid(self,input):
        if (input.isdigit() or input=='') and (len(input)<3):
            return True
        else:
            return False
    def disable_all(self):
        self.cbutton['state']='disabled'
        self.advanbutton['state']='disabled'
        self.fbutton['state']='disabled'
        self.tbutton['state']='disabled'
        if not self.hidden:
            self.radiobuttton1['state']='disabled'
            self.radiobuttton2['state']='disabled'
            self.radiobuttton3['state']='disabled'
            self.fqentry['state']='disabled'
    def enable_all(self):
        self.cbutton['state']='normal'
        self.advanbutton['state']='normal'
        self.fbutton['state']='normal'
        if not self.hidden:
            self.radiobuttton1['state']='normal'
            self.radiobuttton2['state']='normal'
            self.radiobuttton3['state']='normal'
            self.fqentry['state']='normal'
    def date_format(self,date):
        return(str(datetime.datetime.strptime(str(date),"%Y-%m-%d").strftime("%d/%m/%Y")))
if  __name__=='__main__':
    root=tk.Tk()
    root.iconbitmap(default='icon.ico')
    MainApplication(root).pack(side='top',fill='both',expand=True)
    root.mainloop()

