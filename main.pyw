"""
Application to extract data from a table on a mysql database based on a predifned frequency interval
and load this data on an excel spreadsheet.
Credits:

Icon made by VisualPharm at icon-icons.com/icon/database-the-application/2803 (License : CC Attribution)
"""
import tkinter as tk
from tkinter import ttk
from checkboxtreeview import CheckboxTreeview
import datetime
from tkcalendar import Calendar
from xlsscript import ParseData
from sqlscript import GetSignals
from signaltree import Tree,Node
import threading
import pickle
from os import system,path,mkdir,remove
w=1400
h=700
class Preferences(tk.Frame):
    options_default = {
    'Excel':{
        1:0
    },
    'Template':{
        1:''
    }
}
    def __init__(self,parent,MainApp):
        tk.Frame.__init__(self,parent)
        self.mainapp=MainApp

        self.grid_columnconfigure(0,weight=1)
        if path.isfile('settings'):
            with open('settings','rb') as file:
                self.options = pickle.load(file)
        else:
            self.options = Preferences.options_default

        self.excel1var = tk.IntVar(self,self.options['Excel'][1])
        self.excel1var.trace('w',self.column_hide)

        ttk.Separator(self,orient=tk.HORIZONTAL).grid(row=1,column=0,sticky='ew')
        tk.Label(self,text='1) Template Settings').grid(row=1,column=0,sticky='w')
        tk.Label(self,text='Default template').grid(row=2,column=0,sticky='w')
        self.default_template_label = tk.Label(self,text=self.options['Template'][1])
        self.default_template_label.grid(row=2,column=1)
        self.default_template_button = tk.Button(self,text='Set currenly selected template as default',command=self.set_default)
        if not path.isfile(self.mainapp.template_text_box.get('1.0',tk.END).rstrip()):
            self.default_template_button['state']='disabled'
            self.options['Template'][1]=self.default_template_label['text']=''
        self.default_template_button.grid(row=2,column=2)

        ttk.Separator(self,orient=tk.HORIZONTAL).grid(row=3,column=0,sticky='ew')

        tk.Label(self,text='2) Excel Settings').grid(row=3,column=0,sticky='w')
        tk.Label(self,text='Columns hidden in template stay hidden in report.\n(Note that this will prevent Signal Report from creating groups if multiple signals are created)').grid(row=4,column=0)
        self.column_hide_checkbutton = tk.Checkbutton(self,variable=self.excel1var)
        self.column_hide_checkbutton.grid(row=4,column=1)
    def set_default(self):
        self.options['Template'][1] = self.default_template_label['text']=self.mainapp.template_text_box.get('1.0',tk.END)
    def column_hide(self,*args):
        self.options['Excel'][1]=self.excel1var.get()
            
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
                self.parent.advanbtn_text.set("Show Advanced Options")
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
        self.menubar = tk.Menu(self.parent)
        self.menubar.add_command(label='Preferences',command=self.preferencesWindow)
        self.parent.config(menu=self.menubar)
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
        self.timezonemenu.place(x=350,y=23)
        self.advanbutton.place(x=470,y=25)
        self.create_text(15,200,text="Object Path(s): ",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.container.place(x=150,y=200)
        self.listbox.grid()
        self.ysb.grid(row=0, column=1, sticky='ns')
        self.template_text_box.place(x=350,y=560)
        self.template_browse_button.place(x=1050,y=560)
        self.template_clear_button.place(x=1200,y=560)
        self.cbutton.place(x=150,y=560)
    def extract(self):
        def thread_extract():
            self.progress.place(x=500,y=650)
            self.progress.start()
            self.object_fullpaths=self.listbox.get(0,tk.END)
            if path.isfile(self.template_text_box.get('1.0',tk.END).rstrip()):
                self.template_path = self.template_text_box.get('1.0',tk.END).rstrip()
            else:
                self.template_path = None
            if self.hidden:
                extraction=ParseData(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day,int(self.fhour.get()),int(self.fmin.get())),
                datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day,int(self.thour.get()),int(self.tmin.get())),
                self.timezonevar.get(),self.hidden,self.object_fullpaths,self.navbar.optionvar.get(),template_path=self.template_path)
                found=extraction.result
            else:
                if self.navbar.optionvar.get() == 'All Measurements':
                    option = self.optionvar_mv.get()
                else:
                    option = self.optionvar_mt.get()
                extraction=ParseData(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day,int(self.fhour.get()),int(self.fmin.get())),
                datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day,int(self.thour.get()),int(self.tmin.get())),
                self.timezonevar.get(),self.hidden,self.object_fullpaths,self.navbar.optionvar.get(),
                self.fqentry.get(),self.radb.get(),option,self.template_path)
                found=extraction.result       
            self.progress.stop()
            self.progress.place_forget()
            if found==-3:
                tk.messagebox.showerror("Error","Too much data to process. Narrow down your search criteria and try again")
            elif found==-2:
                tk.messagebox.showerror("Error","Sorry,something went wrong.\n"+extraction.errormessage)
            elif found==-1:
                tk.messagebox.showerror("Error","Close the excel file and try again.")
            elif found==0:
                tk.messagebox.showwarning("Info","No data was found for the selected signal(s).")
            elif found==1:
                tk.messagebox.showinfo("Extraction Successful !","You can view the records in the file tables.xlsx")
                system("start EXCEL.EXE \"{}\"".format(extraction.file_path))
            elif found==2:
                tk.messagebox.showwarning("Info","Records for the following data could not be found-\n"+',\n'.join(extraction.not_found))
                system("start EXCEL.EXE \"{}\"".format(extraction.file_path))
        if not self.datetimecheck():
            tk.messagebox.showerror("Date Error","The time interval is invalid. Try Again")
            return
        self.disable_all()        
        threading.Thread(target=thread_extract).start()
        self.enable_all()
        
    def setup(self):
        self.to_datetime = datetime.datetime.now()
        self.from_datetime = self.to_datetime +datetime.timedelta(hours=-1)
        self.tdate = datetime.date(self.to_datetime.year,self.to_datetime.month,self.to_datetime.day)
        self.fdate = datetime.date(self.from_datetime.year,self.from_datetime.month,self.from_datetime.day)
        self.fbutton=ttk.Button(self, text='From', command=self.fgetdate,underline=1)
        vcmd = (self.register(self.onValidate),'%P','%W')
        self.fhour = tk.Spinbox(self,from_=0,to=23,wrap=True,width=2,validatecommand=vcmd,validate='key',name='from_hour')
        self.fmin = tk.Spinbox(self,from_=0,to=59,wrap=True,width=2,validatecommand=vcmd,validate='key',name='from_min')
        self.thour = tk.Spinbox(self,from_=0,to=23,wrap=True,width=2,validatecommand=vcmd,validate='key',name='to_hour')
        self.tmin = tk.Spinbox(self,from_=0,to=59,wrap=True,width=2,validatecommand=vcmd,validate='key',name='to_min')
        self.fhour.delete(0,'end')
        self.fhour.insert(0,self.from_datetime.hour)
        self.fmin.delete(0,'end')
        self.fmin.insert(0,self.from_datetime.minute)
        self.thour.delete(0,'end')
        self.thour.insert(0,self.to_datetime.hour)
        self.tmin.delete(0,'end')
        self.tmin.insert(0,self.to_datetime.minute)
        self.tbutton=ttk.Button(self, text='To', command=self.tgetdate)
        self.fstr=tk.StringVar(self,self.date_format(self.fdate))
        self.flabel=tk.Label(self,textvariable=self.fstr,width=10)
        self.tstr=tk.StringVar(self,self.date_format(self.tdate))
        self.tlabel=tk.Label(self,textvariable=self.tstr,width=10)
        self.photoimage=tk.PhotoImage(file="./img/bgimage.png")
        self.parent.geometry("%dx%d" % (w,h))
        self.parent.title("Create Exel Log File")
        self.advanbtn_text=tk.StringVar(self,value="Show Advanced Options")
        self.advanbutton=tk.Button(self,textvariable=self.advanbtn_text,command=self.advance)
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
        self.container = tk.Frame(self)
        self.listbox = tk.Listbox(self.container,width=85,height=20)
        self.ysb = ttk.Scrollbar(self.container, orient='vertical',command=self.listbox.yview)
        self.listbox.config(yscroll=self.ysb.set)
        self.timezoneOPTIONS = (
            'GMT',
            'GMT+1:00',
            'GMT+2:00',
            'GMT+3:00',
            'GMT+4:00',
            'GMT+5:00',
            'GMT+6:00',
            'GMT+7:00',
            'GMT+8:00',
            'GMT+9:00',
            'GMT+10:00',
            'GMT+11:00',
            'GMT+12:00',
            'GMT-11:00',
            'GMT-10:00',
            'GMT-9:00',
            'GMT-8:00',
            'GMT-7:00',
            'GMT-6:00',
            'GMT-5:00',
            'GMT-4:00',
            'GMT-3:00',
            'GMT-2:00',
            'GMT-1:00',
        )
        if path.isfile('settings'):
            with open('settings','rb') as file:
                self.customization = pickle.load(file)
        else:
            self.customization = Preferences.options_default
        self.timezonevar = tk.StringVar(self,value='GMT+4:00')
        self.timezonemenu = tk.OptionMenu(self,self.timezonevar,*self.timezoneOPTIONS)
        self.template_dir_path = '../../../Templates'
        self.template_text_box = tk.Text(self,width=75,height=1,font=('Helvetica',12,))
        self.template_text_box.tag_config('format1',justify='center')
        self.template_text_box.tag_config('format2',foreground='grey')
        if self.customization['Template'][1]:
            self.template_text_box.insert('end',self.customization['Template'][1])
            self.template_text_box.tag_add('format1','1.0',tk.END)
        else:
            self.template_text_box.insert('end','Select excel template for report (optional).')
            self.template_text_box.tag_add('format2','1.0',tk.END)
        self.template_text_box['state']='disabled'
        self.template_browse_button = tk.Button(self,text='Browse',command=self.file_explore,width=20,height=1)
        self.template_clear_button = tk.Button(self,text='Clear',command=self.clear_template_path,width=20,height=1)
        self.cbutton=tk.Button(self,text="Create Excel Report!",command=self.extract,width=20,height=2)
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL,length=200,  mode='determinate')
    def preferencesWindow(self):
        def save_quit():
            self.customization = pref.options
            with open('settings','wb') as file:
                pickle.dump(self.customization,file)
            top.destroy()
        def revert():
            self.customization = Preferences.options_default
            if path.isfile('settings'):
                remove('settings')
        top = tk.Toplevel()
        top.grab_set()
        top.title('Preferences')
        pref = Preferences(top,self)
        pref.grid(row=0,column=0)
        button1 = tk.Button(top,text='Save and close',command=save_quit)
        button2 = tk.Button(top,text='Discard changes and close',command=top.destroy)
        button3 = tk.Button(top,text='Revert changes to default\n(This will take effect on next launch)',command=revert)
        button1.grid(row=2,column=1)
        button2.grid(row=2,column=2)
        button3.grid(row=2,column=3)

    def fgetdate(self):
        def print_sel():
            self.fdate=self.fcal.selection_get()
            self.tdate = self.fdate
            self.fstr.set(self.date_format(self.fdate))
            self.tstr.set(self.date_format(self.tdate))
            self.ftop.destroy()
        self.ftop = tk.Toplevel(self)
        self.ftop.grab_set()

        self.fcal = Calendar(self.ftop, font="Arial 14", selectmode='day',
                        cursor="hand1",year=self.fdate.year,month=self.fdate.month,day=self.fdate.day)

        self.fcal.pack(fill="both", expand=True)
        ttk.Button(self.ftop, text="Go", command=print_sel).pack()
    def tgetdate(self):
        def print_sel():
            self.tdate=self.tcal.selection_get()
            self.tstr.set(self.date_format(self.tdate))
            self.ttop.destroy()
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
            self.fqentry.config(validate='key',validatecommand=(self.reg1,'%P'))
            self.fqentry.bind('<FocusOut>',self.freq_focusout)
            self.advanbtn_text.set("Hide Advanced Options")
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
            self.advanbtn_text.set("Show Advanced Options")
            self.delete('showtext')
            self.fqentry.place_forget()
            self.radiobutton1.place_forget()
            self.radiobutton2.place_forget()
            self.optionmenu_mv.place_forget()
            self.optionmenu_mt.place_forget()
            self.radiobutton3.place_forget()
        self.hidden = not self.hidden
    def datetimecheck(self):
        if self.fhour.get() and self.fmin.get() and self.thour.get() and self.tmin.get():
            return (datetime.datetime.combine(self.fdate,datetime.time(hour=int(self.fhour.get()),minute=int(self.fmin.get())))
            <(datetime.datetime.combine(self.tdate,datetime.time(hour=int(self.thour.get()),minute=int(self.tmin.get()))))
            <=datetime.datetime.now())
        return False
    def file_explore(self):
        if not path.isdir(self.template_dir_path):
            mkdir(self.template_dir_path)
        template_pathwrapper = tk.filedialog.askopenfile(initialdir=self.template_dir_path, title='Select Report Template',filetypes=(('Excel files','*.xlsx'),))
        if template_pathwrapper:
            self.template_text_box['state']='normal'
            self.template_text_box.delete('1.0',tk.END)
            self.template_text_box.insert('end',template_pathwrapper.name)
            self.template_text_box.tag_add('format1','1.0',tk.END)
            self.template_text_box['state']='disabled'
    def clear_template_path(self):
        self.template_text_box['state']='normal'
        self.template_text_box.delete('1.0',tk.END)
        self.template_text_box.insert('end','Select excel template for report (optional).')
        self.template_text_box.tag_add('format2','1.0',tk.END)
        self.template_text_box['state']='disabled'

    def onValidate(self,P,W):
        called_by=W.split('.')[-1]
        if called_by in ['from_hour','to_hour']:
            if P.isdigit() and int(P) in range(24):
                return True
            elif P=='':
                return True
            else:
                return False
        else:
            if P.isdigit() and int(P) in range(60):
                return True
            elif P=='':
                return True
            else:
                return False
    def freq_valid(self,input):
        if (input.isdigit()) and (len(input)<3):
            return True
        elif input == '':
            return True
        else:
            return False
    def freq_focusout(self,*args):
        if self.fqvar.get() == '':
            self.fqvar.set(1)
    def disable_all(self):
        self.cbutton['state']='disabled'
        self.template_browse_button['state']='disabled'
        self.template_clear_button['state']='disabled'
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
        self.template_browse_button['state']='normal'
        self.template_clear_button['state']='normal'
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
    root.iconbitmap(default='./img/icon.ico')
    MainApplication(root).pack(side='top',fill='both',expand=True)
    root.mainloop()

