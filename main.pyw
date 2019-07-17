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
from functools import partial
from os import system
w=1400
h=700

class Navbar(tk.Frame):
    def __init__(self,parent):
        tk.Frame.__init__(self,parent)
        self.parent=parent
        self.internal_nodes = dict()
        self.tree = CheckboxTreeview(self)
        ysb = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        xsb = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        self.tree.heading('#0', text='Signal tree', anchor='w')
        self.tree.grid(ipadx=100,ipady=100,sticky='e')
        ysb.grid(row=0, column=1, sticky='ns')
        xsb.grid(row=1, column=0, sticky='ew')
        self.OPTIONS=['All Signals','All Controls','All Measurements','All Metering']
        self.optionvar=tk.StringVar(self)
        self.optionvar.set(self.OPTIONS[0])
        self.optionmenu=tk.OptionMenu(self,self.optionvar,*self.OPTIONS,command=self.callback)
        self.optionmenu.grid(row=0,column=2)
        self.layout=Tree(GetSignals().result)
        self.all_electrical=self.layout.root['All']['site']
        self.all_system=self.layout.root['All']['scs']
        self.control=self.layout.root['Control']
        self.measurement=self.layout.root['Measurement']
        self.meter=self.layout.root['Meter']
        self.root_iid = []
        self.tree.bind('<<TreeviewSelect>>',self.getchecked)
        self.build_tree('', self.all_electrical,self.all_electrical.data,root=True)
        self.build_tree('', self.all_system,self.all_system.data,root=True)
    def getchecked(self,*args):
        self.parent.listbox.delete(0,tk.END)
        for iid in self.tree.get_checked():
            if not self.tree.get_children(iid):
                self.parent.listbox.insert(tk.END,self.tree.item(iid,'values')[0])
    def build_tree(self,parent_iid,node,data,root=False):
        node_iid=self.tree.insert(parent_iid,"end",text=data)
        if not node.isInternalNode():
            self.tree.item(node_iid,values=[node.absolute_path])
        if root:
            self.root_iid.append(node_iid)
        for child in node.getChildren():
            self.build_tree(node_iid,child,child.data)
    def callback(self,*args):
        self.tree.delete(*self.root_iid)
        self.parent.listbox.delete(0,tk.END)
        self.root_iid=[]
        self.parent.optionmenu_mt.place_forget()
        self.parent.optionmenu_mv.place_forget()
        if self.optionvar.get() in self.OPTIONS[:-2]:
            self.parent.advanbutton['state']='disabled'
            if not self.parent.hidden:
                self.parent.btn_text.set("Show Advanced Options")
                self.parent.delete('showtext')
                self.parent.fqentry.place_forget()
                self.parent.radiobutton1.place_forget()
                self.parent.radiobutton2.place_forget()
                self.parent.radiobutton3.place_forget()
                self.parent.hidden=True  
            if self.optionvar.get() == 'All Controls':
                self.build_tree('', self.control,self.control.data,root=True)
            elif self.optionvar.get() == 'All Signals':
                self.build_tree('', self.all_electrical,self.all_electrical.data,root=True)
                self.build_tree('', self.all_system,self.all_system.data,root=True)
            return
        elif self.optionvar.get() == 'All Measurements':
            if not self.parent.hidden:
                self.parent.optionmenu_mv.place(x=65,y=145)
            self.build_tree('',self.measurement,self.measurement.data,root=True)
        elif self.optionvar.get()=='All Metering':
            if not self.parent.hidden:
                self.parent.optionmenu_mt.place(x=65,y=145)
            self.build_tree('', self.meter,self.meter.data,root=True)
        
        self.parent.advanbutton['state']='normal'


class MainApplication(tk.Canvas):
    def __init__(self,parent):
        tk.Canvas.__init__(self,parent)
        self.parent=parent
        self.setup()
        self.navbar=Navbar(self)
        self.navbar.place(x=700,y=80)
        self.create_image(0,0,image=self.photoimage,anchor='nw')
        self.create_text(15,20,text="Time Period :",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.fbutton.place(x=150,y=20)
        self.tbutton.place(x=250,y=20)
        self.flabel.place(x=150,y=45)
        self.fhour.place(x=155,y=70)
        self.create_text(185,70,text=":",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.fmin.place(x=195,y=70)
        self.tlabel.place(x=250,y=45)
        self.thour.place(x=255,y=70)
        self.create_text(285,70,text=":",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.tmin.place(x=295,y=70)
        self.advanbutton.place(x=350,y=25)
        self.create_text(15,200,text="Object Path(s): ",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.container.place(x=150,y=200)
        self.listbox.grid()
        self.ysb.grid(row=0, column=1, sticky='ns')
        self.cbutton.place(x=150,y=560)
    def extract(self):
        def thread_extract():
            self.progress.place(x=200,y=300)
            self.progress.start()
            self.object_fullpaths=self.listbox.get(0,tk.END)
            if self.hidden:
                extraction=ParseData(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day,int(self.fhourstr.get()),int(self.fminstr.get())),
                datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day,int(self.thourstr.get()),int(self.tminstr.get())),
                self.hidden,self.object_fullpaths,self.navbar.optionvar.get())
                found=extraction.result
            else:
                extraction=ParseData(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day,int(self.fhourstr.get()),int(self.fminstr.get())),
                datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day,int(self.thourstr.get()),int(self.tminstr.get())),
                self.hidden,self.object_fullpath,self.navbar.optionvar.get(),
                self.fqentry.get(),self.radb.get(),self.optionvar.get())
                found=extraction.result       
            self.progress.stop()
            self.progress.place_forget()
            self.enable_all()
            if found==-3:
                tk.messagebox.showerror("Error","Too much data to process. Narrow down your search criteria and try again")
            elif found==-2:
                tk.messagebox.showerror("Error","Something went wrong. Restart the application and try again.")
            elif found==-1:
                tk.messagebox.showerror("Error","Close the excel file and try again.")
            elif found==0:
                tk.messagebox.showwarning("Info","No data was found for the selected signal(s).")
            elif found==1:
                tk.messagebox.showinfo("Extraction Successful !","You can view the records in the file tables.xlsx")
                system("start EXCEL.EXE \"./SignalLog/"+extraction.str_time+".xlsx\"")
            elif found==2:
                tk.messagebox.showwarning("Info","Records for the following data could not be found-\n"+','.join([extraction.not_found]))
        self.disable_all()        
        threading.Thread(target=thread_extract).start()
        
    def setup(self):
        self.tdate = calendar.datetime.date.today()
        self.fdate=calendar.datetime.date(self.tdate.year,self.tdate.month,self.tdate.day)+datetime.timedelta(days=-2)
        self.fbutton=ttk.Button(self, text='From', command=self.fgetdate,underline=1)
        self.fhourstr=tk.StringVar(self,str((datetime.datetime.now()+datetime.timedelta(hours=-2)).hour))
        self.fhour = tk.Spinbox(self,from_=0,to=23,wrap=True,textvariable=self.fhourstr,width=2,state='readonly')
        self.fminstr=tk.StringVar(self,'30')
        self.fmin = tk.Spinbox(self,from_=0,to=59,wrap=True,textvariable=self.fminstr,width=2,state='readonly')
        self.thourstr=tk.StringVar(self,str((datetime.datetime.now()+datetime.timedelta(hours=-1)).hour))
        self.thour = tk.Spinbox(self,from_=0,to=23,wrap=True,textvariable=self.thourstr,width=2,state='readonly')
        self.tminstr=tk.StringVar(self,'30')
        self.tmin = tk.Spinbox(self,from_=0,to=59,wrap=True,textvariable=self.tminstr,width=2,state='readonly')
        self.prev_valid = [self.fhourstr.get(),self.fminstr.get(),self.thourstr.get(),self.tminstr.get()]
        self.tbutton=ttk.Button(self, text='To', command=self.tgetdate)
        self.fstr=tk.StringVar(self,self.date_format(self.fdate))
        self.flabel=tk.Label(self,textvariable=self.fstr,width=10)
        self.tstr=tk.StringVar(self,self.date_format(self.tdate))
        self.tlabel=tk.Label(self,textvariable=self.tstr,width=10)
        self.photoimage=tk.PhotoImage(file="C:\\Users\\OISM\\Desktop\\sqlApp\\bgimage.png")
        self.parent.geometry("%dx%d" % (w,h))
        self.parent.title("Create Exel Log File")
        self.btn_text=tk.StringVar(self,value="Show Advanced Options")
        self.advanbutton=tk.Button(self,textvariable=self.btn_text,command=self.advance)
        self.advanbutton['state']='disabled'
        self.hidden=True
        self.fqvar=tk.StringVar(self,value='1')
        self.radb=tk.IntVar(self,2)
        self.OPTIONS_MEASUREMENT=['Changes','Average']
        self.OPTIONS_METERING = ['Changes','Consumption']
        self.optionvar_mv=tk.StringVar(self)
        self.optionvar_mv.set(self.OPTIONS_MEASUREMENT[0])
        self.optionmenu_mv=tk.OptionMenu(self,self.optionvar_mv,*self.OPTIONS_MEASUREMENT)
        self.optionmenu_mv.config(width=10)
        self.optionvar_mt = tk.StringVar(self)
        self.optionvar_mt.set(self.OPTIONS_METERING[0])
        self.optionmenu_mt =tk.OptionMenu(self,self.optionvar_mt,*self.OPTIONS_METERING)
        self.optionmenu_mt.config(width=10)
        self.flag=tk.IntVar(self,0)
        self.from_within = 0
        self.flag.trace('w',self.flag_callback)
        self.fhourstr.trace('w',self.time_callback)
        self.fminstr.trace('w',self.time_callback)
        self.thourstr.trace('w',self.time_callback)
        self.tminstr.trace('w',self.time_callback)
        self.container = tk.Frame(self)
        self.listbox = tk.Listbox(self.container,width=85,height=20)
        self.ysb = ttk.Scrollbar(self.container, orient='vertical',command=self.listbox.yview)
        self.listbox.config(yscroll=self.ysb.set)
        self.cbutton=tk.Button(self,text="Create excel File!",command=self.extract,width=20,height=2)
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL,length=100,  mode='indeterminate')
    def fgetdate(self):
        def print_sel():
            if self.datetimecheck(fdate=self.fcal.selection_get()):
                self.flag.set(0)
                self.fdate=self.fcal.selection_get()
                self.tdate = self.fdate
                self.fstr.set(self.date_format(self.fdate))
                self.tstr.set(self.date_format(self.tdate))
                self.ftop.destroy()
            else:
                self.flag.set(1)
        self.ftop = tk.Toplevel(self)
        self.ftop.grab_set()

        self.fcal = Calendar(self.ftop, font="Arial 14", selectmode='day',
                        cursor="hand1",year=self.fdate.year,month=self.fdate.month,day=self.fdate.day)

        self.fcal.pack(fill="both", expand=True)
        ttk.Button(self.ftop, text="Go", command=print_sel).pack()
    def tgetdate(self):
        def print_sel():
            if self.datetimecheck(tdate=self.tcal.selection_get()):
                self.flag.set(0)
                self.tdate=self.tcal.selection_get()
                self.tstr.set(self.date_format(self.tdate))
                self.ttop.destroy()
            else:
                self.flag.set(1)
        self.ttop = tk.Toplevel(self)
        self.ttop.grab_set()
        self.tcal = Calendar(self.ttop, font="Arial 14", selectmode='day',
                        cursor="hand1",year=self.fdate.year,month=self.fdate.month,day=self.fdate.day)
        self.tcal.pack(fill="both", expand=True)
        ttk.Button(self.ttop, text="Go", command=print_sel).pack()
    def advance(self):
        if self.hidden:
            self.radiobutton1=tk.Radiobutton(self,text="Second(s)",value=1,variable=self.radb)
            self.radiobutton2=tk.Radiobutton(self,text="Minute(s)",value=2,variable=self.radb)
            self.radiobutton3=tk.Radiobutton(self,text="Hour(s)",value=3,variable=self.radb)
            self.fqentry=tk.Entry(self,width=2,textvariable=self.fqvar)
            self.reg1=self.register(self.freq_valid)
            self.fqentry.config(validate='focusout',validatecommand=(self.reg1,'%P'),invalidcommand=self.freq_invalid)
            self.btn_text.set("Hide Advanced Options")
            self.create_text(15,150,text="Show\t\t\tevery".expandtabs(11),fill='white',anchor='nw',font=("Arial",12,'bold'),tag='showtext')
            if self.navbar.optionvar.get() == 'All Measurements':
                self.optionmenu_mv.place(x=65,y=145)
            else:
                self.optionmenu_mt.place(x=65,y=145) 
            self.fqentry.place(x=230,y=153)
            self.radiobutton1.place(x=260,y=150)
            self.radiobutton2.place(x=340,y=150)
            self.radiobutton3.place(x=420,y=150)
        else:
            self.btn_text.set("Show Advanced Options")
            self.delete('showtext')
            self.fqentry.place_forget()
            self.radiobutton1.place_forget()
            self.radiobutton2.place_forget()
            self.optionmenu_mv.place_forget()
            self.optionmenu_mt.place_forget()
            self.radiobutton3.place_forget()
        self.hidden = not self.hidden
    def datetimecheck(self,fdate=None,tdate=None):
        fhour = int(self.fhourstr.get())
        fmin = int(self.fminstr.get())
        thour = int(self.thourstr.get())
        tmin = int(self.tminstr.get())
        if fdate:
            return (datetime.datetime.combine(fdate,datetime.time(hour=fhour,minute=fmin))<
            datetime.datetime.combine(self.tdate,datetime.time(hour=thour,minute=tmin)))
        elif tdate:
            return (datetime.datetime.combine(self.fdate,datetime.time(hour=fhour,minute=fmin))<
            datetime.datetime.combine(tdate,datetime.time(hour=thour,minute=tmin)) and datetime.datetime.combine(tdate,datetime.time(hour=thour,minute=tmin))
            <=datetime.datetime.now())
        else:
            return (datetime.datetime.combine(self.fdate,datetime.time(hour=fhour,minute=fmin))<
            datetime.datetime.combine(self.tdate,datetime.time(hour=thour,minute=tmin)) and datetime.datetime.combine(self.tdate,datetime.time(hour=thour,minute=tmin))
            <=datetime.datetime.now())
    def flag_callback(self,*args):
        if self.flag.get() == 1:
            messagebox.showerror("Date Error","The time interval is invalid. Try Again")
    def time_callback(self,*agrs):
        if not self.from_within:
            if not self.datetimecheck():
                self.flag.set(1)
                self.from_within=1
                self.fhourstr.set(self.prev_valid[0])
                self.fminstr.set(self.prev_valid[1])
                self.thourstr.set(self.prev_valid[2])
                self.tminstr.set(self.prev_valid[3])
                self.from_within=0
                self.flag.set(0)
            else:
                self.prev_valid=[self.fhourstr.get(),self.fminstr.get(),self.thourstr.get(),self.tminstr.get()]
    def freq_invalid(self):
        self.fqvar.set('1')
    def freq_valid(self,input):
        if (input.isdigit()) and (len(input)<3):
            valid = True
        else:
            valid = False
        if not valid:
            self.fqentry.after_idle(lambda: self.fqentry.config(validate='focusout'))
        return valid
    def disable_all(self):
        self.cbutton['state']='disabled'
        self.advanbutton['state']='disabled'
        self.fbutton['state']='disabled'
        self.tbutton['state']='disabled'
        if not self.hidden:
            self.radiobutton1['state']='disabled'
            self.radiobutton2['state']='disabled'
            self.radiobutton3['state']='disabled'
            self.fqentry['state']='disabled'
            self.optionmenu_mv['state']='disabled'
    def enable_all(self):
        self.cbutton['state']='normal'
        self.advanbutton['state']='normal'
        self.fbutton['state']='normal'
        if not self.hidden:
            self.radiobutton1['state']='normal'
            self.radiobutton2['state']='normal'
            self.radiobutton3['state']='normal'
            self.fqentry['state']='normal'
            self.optionmenu_mv['state']='normal'
    def date_format(self,date):
        return(str(datetime.datetime.strptime(str(date),"%Y-%m-%d").strftime("%d/%m/%Y")))
if  __name__=='__main__':
    root=tk.Tk()
    root.iconbitmap(default='icon.ico')
    MainApplication(root).pack(side='top',fill='both',expand=True)
    root.mainloop()

