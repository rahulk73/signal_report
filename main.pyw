"""
Credits:
Icon made by VisualPharm at icon-icons.com/icon/database-the-application/2803 (License : CC Attribution)
Background Image by <a href="https://pixabay.com/users/Clker-Free-Vector-Images-3736/?utm_source=link-attribution&amp;utm_medium=referral&amp;utm_campaign=image&amp;utm_content=34536">Clker-Free-Vector-Images</a> from <a href="https://pixabay.com/?utm_source=link-attribution&amp;utm_medium=referral&amp;utm_campaign=image&amp;utm_content=34536">Pixabay</a>

"""
import tkinter as tk
from tkinter import ttk
from checkboxtreeview import CheckboxTreeview
import datetime
from tkcalendar import Calendar
from xlsscript import AnalogReport,Preferences,EventReport
from sqlscript import GetSignals,AccessDeniedError
from signaltree import Tree,Node
import threading
import pickle
from os import system,path,mkdir,remove

class ExcelPreferences(tk.Frame):
    def __init__(self,parent,MainApp,options):
        tk.Frame.__init__(self,parent)
        self.grid_columnconfigure(0,weight=1)
        self.parent = parent
        self.options = options
        self.mainapp = MainApp
        self.column_hide_var = tk.IntVar(self,self.options[1])
        self.column_hide_var.trace('w',self.column_hide)
        self.seperate_worksheet_var = tk.IntVar(self,self.options[2])
        self.seperate_worksheet_var.trace('w',self.seperate_worksheet)
        self.MODES = [
            ('Area Chart','1'),
            ('Verical Bar Chart','2'),
            ('Horizontal Bar Chart','3'),
            ('Line Chart','4'),
            ('None','0')

        ]
        self.chart_var = tk.StringVar(self,str(self.options[3]))
        self.chart_var.trace('w',self.set_chart)

        tk.Label(self,text='Columns hidden in template stay hidden in report.').grid(row=0,column=0,sticky='w')
        self.column_hide_checkbutton = tk.Checkbutton(self,variable=self.column_hide_var)
        self.column_hide_checkbutton.grid(row=0,column=1)

        tk.Label(self,text='Always use a seperate worksheet for each signal.').grid(row=1,column=0,sticky='w')
        self.seperate_worksheet_checkbutton = tk.Checkbutton(self,variable=self.seperate_worksheet_var)
        self.seperate_worksheet_checkbutton.grid(row=1,column=1)

        tk.Label(self,text='Default chart').grid(row=2,column=0,sticky='w',rowspan=2)
        for i,mode in enumerate(self.MODES,start=1):
                tk.Radiobutton(self,text=mode[0],variable=self.chart_var,value=mode[1]).grid(row=2,column=i)
    def column_hide(self,*args):
        self.options[1]=self.column_hide_var.get()
    def seperate_worksheet(self,*args):
        self.options[2]=self.seperate_worksheet_var.get()
    def set_chart(self,*args):
        self.options[3]=int(self.chart_var.get())

class TemplatePreferences(tk.Frame):
    def __init__(self,parent,MainApp,options):
        tk.Frame.__init__(self,parent)
        self.grid_columnconfigure(0,weight=1)
        self.parent = parent
        self.options = options
        self.mainapp = MainApp

        tk.Label(self,text='Default template').grid(row=0,column=0,sticky='w')
        self.default_template_label = tk.Label(self,text=self.options[1])
        self.default_template_label.grid(row=0,column=1,sticky='w')
        self.default_template_button = tk.Button(self,text='Set currenly selected template as default',command=self.set_default_template)
        if not path.isfile(self.mainapp.template_text_box.get('1.0',tk.END).rstrip()):
            self.default_template_button['state']='disabled'
            self.options[1]=self.default_template_label['text']=''
        self.default_template_button.grid(row=0,column=2,sticky='w')

    def set_default_template(self):
        self.options[1] = self.default_template_label['text']=self.mainapp.template_text_box.get('1.0',tk.END)
class MySQLPreferences(tk.Frame):
    def __init__(self,parent,MainApp,options):
        tk.Frame.__init__(self,parent)
        self.grid_columnconfigure(0,weight=1)
        self.parent = parent
        self.options = options
        self.mainapp = MainApp

        tk.Label(self,text='Server Host').grid(row=0,column=0,sticky='w')
        self.host_var = tk.StringVar(self,self.options['sh'])
        self.host_var.trace('w',self.callback)
        tk.Entry(self,textvariable=self.host_var).grid(row=0,column=1,sticky='e')

        tk.Label(self,text='Username').grid(row=1,column=0,sticky='w')
        self.un_var = tk.StringVar(self,self.options['un'])
        self.un_var.trace('w',self.callback2)
        tk.Entry(self,textvariable=self.un_var).grid(row=1,column=1,sticky='e')

        tk.Label(self,text='Password').grid(row=2,column=0,sticky='w')
        self.pw_var = tk.StringVar(self,self.options['pw'])
        self.pw_var.trace('w',self.callback3)
        tk.Entry(self,textvariable=self.pw_var,show='*').grid(row=2,column=1,sticky='e')

        tk.Label(self,text='Default Schema').grid(row=3,column=0,sticky='w')
        self.db_var = tk.StringVar(self,self.options['db'])
        self.db_var.trace('w',self.callback4)
        tk.Entry(self,textvariable=self.db_var).grid(row=3,column=1,sticky='e')

    def callback(self,*args):
        self.options['sh'] = self.host_var.get()
    def callback2(self,*args):
        self.options['un'] = self.un_var.get()
    def callback3(self,*args):
        self.options['pw'] = self.pw_var.get()
    def callback4(self,*args):
        self.options['db'] = self.db_var.get()


class PreferencesContainer(tk.Frame):
    def __init__(self,parent,MainApp):
        tk.Frame.__init__(self,parent)
        self.mainapp=MainApp
        self.grid_columnconfigure(0,weight=1)
        if path.isfile('settings'):
            with open('settings','rb') as file:
                self.options = pickle.load(file)
        else:
            self.options = Preferences.options_default
            
        ttk.Separator(self,orient=tk.HORIZONTAL).grid(row=1,column=0,sticky='ew')
        tk.Label(self,text='Excel Settings',font='Helvetica 12 bold').grid(row=1,column=0,sticky='ns')
        ExcelPreferences(self,self.mainapp,self.options['Excel']).grid(sticky='ew',row=2,column=0)

        ttk.Separator(self,orient=tk.HORIZONTAL).grid(row=3,column=0,sticky='ew')
        tk.Label(self,text='Template Settings',font=('Helvetica 12 bold')).grid(row=3,column=0,sticky='ns')
        TemplatePreferences(self,self.mainapp,self.options['Template']).grid(sticky='ew',row=4,column=0)

        ttk.Separator(self,orient=tk.HORIZONTAL).grid(row=5,column=0,sticky='ew')
        tk.Label(self,text='MySQL Settings',font=('Helvetica 12 bold')).grid(row=5,column=0,sticky='ns')
        MySQLPreferences(self,self.mainapp,self.options['MySQL']).grid(sticky='ew',row=6,column=0)
class Navbar(tk.Frame):
    def __init__(self,parent):
        tk.Frame.__init__(self,parent)
        self.parent=parent
        self.internal_nodes = dict()
        self.tree = CheckboxTreeview(self)
        ysb = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        xsb = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        self.tree.heading('#0', text='Signal Tree', anchor='w')
        self.tree.grid(ipadx=100,ipady=100,sticky='e')
        ysb.grid(row=0, column=1, sticky='ns')
        xsb.grid(row=1, column=0, sticky='ew')
        self.OPTIONS=['All Signals','All Controls','All Measurements','All Metering']
        self.optionvar=tk.StringVar(self)
        self.optionvar.set(self.OPTIONS[2])
        self.optionmenu=tk.OptionMenu(self,self.optionvar,*self.OPTIONS[-2:],command=self.callback)
        self.optionmenu.grid(row=0,column=2)
        self.layout=Tree(GetSignals(**self.parent.customization['MySQL']).result)
        self.all_electrical=self.layout.root['All']['site']
        self.all_system=self.layout.root['All']['scs']
        self.control=self.layout.root['Control']
        self.measurement=self.layout.root['Measurement']
        self.meter=self.layout.root['Meter']
        self.root_iid = []
        self.tree.bind('<<TreeviewSelect>>',self.getchecked)
        self.build_tree('',self.measurement,self.measurement.data,root=True)
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
            self.parent.disable_advance()
            if self.optionvar.get() == 'All Controls':
                self.build_tree('', self.control,self.control.data,root=True)
            elif self.optionvar.get() == 'All Signals':
                self.build_tree('', self.all_electrical,self.all_electrical.data,root=True)
                self.build_tree('', self.all_system,self.all_system.data,root=True)
            return
        elif self.optionvar.get() == 'All Measurements':
            if not self.parent.hidden:
                self.parent.optionmenu_mv.place(x=65,y=325)
            self.build_tree('',self.measurement,self.measurement.data,root=True)
        elif self.optionvar.get()=='All Metering':
            if not self.parent.hidden:
                self.parent.optionmenu_mt.place(x=65,y=325)
            self.build_tree('', self.meter,self.meter.data,root=True)
        
        self.parent.advanbutton['state']='normal'
    def report_change(self,report_type):
        if report_type == self.parent.reportOPTIONS[0]:
            self.optionvar.set(self.OPTIONS[2])
            self.optionmenu=tk.OptionMenu(self,self.optionvar,*self.OPTIONS[-2:],command=self.callback)
        else:
            self.optionvar.set(self.OPTIONS[0])
            self.optionmenu=tk.OptionMenu(self,self.optionvar,*self.OPTIONS[:-2],command=self.callback)
        self.optionmenu.grid(row=0,column=2)
        self.callback()


class MainApplication(tk.Canvas):
    def __init__(self,parent):
        tk.Canvas.__init__(self,parent)
        self.parent=parent
        self.setup()
        self.menubar = tk.Menu(self.parent)
        self.menubar.add_command(label='Preferences',command=self.preferencesWindow)
        self.menubar.add_command(label='About',command=self.aboutWindow)
        self.parent.config(menu=self.menubar)
        try:
            self.navbar=Navbar(self)
            self.navbar.place(x=700,y=260)
        except AccessDeniedError:
            tk.messagebox.showerror("Access to MySQL Databse denied","Unable to connect to mysql databse with current credentials.")
            tk.Label(self,text="Not connected to database.\nSelect 'Preferences', enter the appropriate credentials, select 'Save & Close' and restart the program.").place(x=800,y=300)
        self.create_image(0,0,image=self.photoimage,anchor='nw')
        self.create_text(15,200,text="Time Period :",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.fbutton.place(x=150,y=200)
        self.tbutton.place(x=250,y=200)
        self.flabel.place(x=150,y=225)
        self.fhour.place(x=155,y=250)
        self.create_text(185,250,text=":",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.fmin.place(x=195,y=250)
        self.tlabel.place(x=250,y=225)
        self.thour.place(x=255,y=250)
        self.create_text(285,250,text=":",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.tmin.place(x=295,y=250)
        self.timezonemenu.place(x=490,y=203)
        self.report_type_menu.place(x=350,y=203)
        self.advanbutton.place(x=610,y=205)
        self.create_text(15,380,text="Object Path(s): ",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.container.place(x=150,y=380)
        self.listbox.grid()
        self.ysb.grid(row=0, column=1, sticky='ns')
        self.template_text_box.place(x=350,y=750)
        self.template_browse_button.place(x=1050,y=750)
        self.template_clear_button.place(x=1200,y=750)
        self.cbutton.place(x=150,y=740)
    def extract(self):
        def thread_extract():
            self.progress.place(x=500,y=830)
            self.progress.start()
            self.object_fullpaths=self.listbox.get(0,tk.END)
            if path.isfile(self.template_text_box.get('1.0',tk.END).rstrip()):
                self.template_path = self.template_text_box.get('1.0',tk.END).rstrip()
            else:
                self.template_path = None
            if self.report_type_var.get() == self.reportOPTIONS[1] and len(self.object_fullpaths)<1:
                if self.event_duration_var.get() != self.eventdurationOPTIONS[-1]:
                    time_unit = dict(zip(self.eventdurationOPTIONS[:-1],[
                        60,
                        60*60,
                        24*60*60,
                        7*24*60*60,
                        30*24*60*60,
                        365*24*60*60
                    ]))
                    seconds = -1*int(self.fqentry2.get())*(time_unit.get(self.event_duration_var.get()))
                    fdate = datetime.datetime.now() + datetime.timedelta(seconds=seconds)
                    tdate = datetime.datetime.now()
                else:
                    fdate = datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day,int(self.fhour.get()),int(self.fmin.get()))
                    tdate =  datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day,int(self.thour.get()),int(self.tmin.get()))
                report = EventReport(
                    fdate,
                    tdate,
                    self.timezonevar.get(),
                    self.template_path,
                )


            elif self.hidden:
                report = AnalogReport(
                    datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day,int(self.fhour.get()),int(self.fmin.get())),
                    datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day,int(self.thour.get()),int(self.tmin.get())),
                    self.timezonevar.get(),
                    self.hidden,
                    self.object_fullpaths,
                    self.navbar.optionvar.get(),
                    template_path=self.template_path,           
                )
            else:
                if self.navbar.optionvar.get() == 'All Measurements':
                    data_type = self.optionvar_mv.get()
                else:
                    data_type = self.optionvar_mt.get()
                report = AnalogReport(
                    datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day,int(self.fhour.get()),int(self.fmin.get())),
                    datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day,int(self.thour.get()),int(self.tmin.get())),
                    self.timezonevar.get(),
                    self.hidden,
                    self.object_fullpaths,
                    self.navbar.optionvar.get(),
                    self.fqentry.get(),
                    self.radb.get(),
                    data_type,
                    self.template_path
                )

            found=report.result       
            self.progress.stop()
            self.progress.place_forget()
            self.configure_all(state='normal')
            if found==-4:
                tk.messagebox.showerror("Error","Sorry,something went wrong.\n"+report.errormessage)
            elif found==-3:
                tk.messagebox.showerror("Error","Too much data to process. Narrow down your search criteria and try again")
            elif found==-2:
                tk.messagebox.showerror("Permission Error","If the template file is open, close it and try again.\nTry running the programme with Admin priviliges.")
            elif found==-1:
                tk.messagebox.showerror("Access to MySQL Databse denied","Unable to connect to mysql databse with current credentials.")
            elif found==0:
                tk.messagebox.showwarning("Info","No data was found for the selected signal(s).")
            elif found==1:
                tk.messagebox.showinfo("Extraction Successful !","You can view the records in the file tables.xlsx")
                system("start EXCEL.EXE \"{}\"".format(report.file_path))
            elif found==2:
                tk.messagebox.showwarning("Info","Records for the following data could not be found-\n"+',\n'.join(report.workbook.not_found))
                system("start EXCEL.EXE \"{}\"".format(report.file_path))
        if not self.datetimecheck():
            tk.messagebox.showerror("Date Error","The time interval is invalid. Try Again")
            return
        self.configure_all(state='disabled')        
        threading.Thread(target=thread_extract).start()
        
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
        self.photoimage=tk.PhotoImage(file="./img/bgimage2.png")
        self.resize_dim = (1400,700)
        self.parent.geometry("{0}x{1}+0+0".format(self.parent.winfo_screenwidth(),self.parent.winfo_screenheight()))
        self.parent.title("Signal Report")
        self.advanbtn_text=tk.StringVar(self,value="Show Advanced Options")
        self.advanbutton=tk.Button(self,textvariable=self.advanbtn_text,command=self.advance)
        self.hidden=True
        self.fqvar=tk.StringVar(self,value='1')
        self.fqentry=tk.Entry(self,width=2,textvariable=self.fqvar)
        self.reg1=self.register(self.freq_valid)
        self.fqentry.config(validate='key',validatecommand=(self.reg1,'%P'))
        self.fqentry.bind('<FocusOut>',self.freq_focusout)
        self.fqvar2=tk.StringVar(self,value='1')
        self.fqentry2=tk.Entry(self,width=2,textvariable=self.fqvar)
        self.reg1=self.register(self.freq_valid)
        self.fqentry2.config(validate='key',validatecommand=(self.reg1,'%P'))
        self.fqentry2.bind('<FocusOut>',self.freq_focusout)
        self.radb=tk.IntVar(self,2)
        self.radiobutton1=tk.Radiobutton(self,text="Second(s)",value=1,variable=self.radb)
        self.radiobutton2=tk.Radiobutton(self,text="Minute(s)",value=2,variable=self.radb)
        self.radiobutton3=tk.Radiobutton(self,text="Hour(s)",value=3,variable=self.radb)
        self.radiobutton4=tk.Radiobutton(self,text="Day(s)",value=4,variable=self.radb)
        self.radiobutton5=tk.Radiobutton(self,text="Week(s)",value=5,variable=self.radb)
        self.radiobutton6=tk.Radiobutton(self,text="Month(s)",value=6,variable=self.radb)
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
        self.template_dir_path = 'Templates'
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
        self.reportOPTIONS = (
            'Analog Report',
            'Event Report',
        )
        self.eventdurationOPTIONS = (
            'Minute(s)',
            'Hour(s)',
            'Day(s)',
            'Week(s)',
            'Month(s)',
            'Year(s)',
            'Custom',
        )
        self.event_duration_var = tk.StringVar(self,value=self.eventdurationOPTIONS[1])
        self.event_duration_var.trace('w',self.eventDurationCallback)
        self.event_duration_menu = tk.OptionMenu(self,self.event_duration_var,*self.eventdurationOPTIONS)
        self.event_duration_menu.config(width=10)
        self.report_type_var = tk.StringVar(self,value=self.reportOPTIONS[0])
        self.prev_report_type = self.report_type_var.get()
        self.report_type_var.trace('w',self.changeReportScreen)
        self.report_type_menu = tk.OptionMenu(self,self.report_type_var,*self.reportOPTIONS)
        self.cbutton=tk.Button(self,text="Create Report!",command=self.extract,width=20,height=2)
        self.progress = ttk.Progressbar(self, orient=tk.HORIZONTAL,length=200,  mode='determinate')
        if not path.isdir('./Templates'):
            mkdir('./Templates')
    def changeReportScreen(self,*args):
        if self.prev_report_type == self.report_type_var.get():
            return
        {self.reportOPTIONS[0]:self.changeToEvent,self.reportOPTIONS[1]:self.changeToAnalog}.get(self.prev_report_type)()
        self.navbar.report_change(self.report_type_var.get())
    def changeToAnalog(self):
        self.delete('event_text')
        self.fqentry2.place_forget()
        self.event_duration_menu.place_forget()
        self.event_duration_var.set(self.eventdurationOPTIONS[-1])

        self.prev_report_type = self.reportOPTIONS[0]
    def changeToEvent(self):
        self.create_text(150,285,text="Last",fill='white',anchor='nw',font=("Arial",12,'bold'),tag='event_text')
        self.fqentry2.place(x=200,y=285)
        self.event_duration_menu.place(x=225,y=280)
        self.disable_advance()
        self.event_duration_var.set(self.eventdurationOPTIONS[1])

        self.prev_report_type = self.reportOPTIONS[1]
    def eventDurationCallback(self,*args):
        if self.event_duration_var.get() == self.eventdurationOPTIONS[-1]:
            state = 'normal'
            self.delete('event_text')
            self.fqentry2.place_forget()
        else:
            state='disabled'
            self.fqentry2.place(x=200,y=285)
            self.create_text(150,285,text="Last",fill='white',anchor='nw',font=("Arial",12,'bold'),tag='event_text')
        self.fbutton['state']=state
        self.tbutton['state']=state
        self.fhour.config(state=state)
        self.thour.config(state=state)
        self.fmin.config(state=state)
        self.tmin.config(state=state)
        

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
        pref = PreferencesContainer(top,self)
        pref.grid(row=0,column=0)
        button1 = tk.Button(top,text='Save and close',command=save_quit)
        button2 = tk.Button(top,text='Discard changes and close',command=top.destroy)
        button3 = tk.Button(top,text='Revert changes to default\n(This will take effect on next launch)',command=revert)
        button1.grid(row=2,column=1)
        button2.grid(row=2,column=2)
        button3.grid(row=2,column=3)
    def aboutWindow(self):
        self.about_message = """Version 0.15.11\n 
Commit babbee69f433fef81ff94e7453a5dcc3475b9377\n
Signal Report © (All Rights Reserved) is an open source project that was created by Farhan Ali, Arun Aery and Rahul Kumar at Schneider Electric Dubai.
        """
        tk.messagebox.showinfo("About",self.about_message)

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
            self.advanbtn_text.set("Hide Advanced Options")
            self.create_text(15,330,text="Show\t\t\tevery".expandtabs(11),fill='white',anchor='nw',font=("Arial",12,'bold'),tag='showtext')
            if self.navbar.optionvar.get() == 'All Measurements':
                self.optionmenu_mv.place(x=65,y=325)
            else:
                self.optionmenu_mt.place(x=65,y=325) 
            self.fqentry.place(x=230,y=333)
            self.radiobutton1.place(x=260,y=330)
            self.radiobutton2.place(x=340,y=330)
            self.radiobutton3.place(x=415,y=330)
            self.radiobutton4.place(x=480,y=330)
            self.radiobutton5.place(x=540,y=330)
            self.radiobutton6.place(x=605,y=330)
        else:
            self.advanbtn_text.set("Show Advanced Options")
            self.delete('showtext')
            self.fqentry.place_forget()
            self.radiobutton1.place_forget()
            self.radiobutton2.place_forget()
            self.radiobutton3.place_forget()
            self.radiobutton4.place_forget()
            self.radiobutton5.place_forget()
            self.radiobutton6.place_forget()
            self.optionmenu_mv.place_forget()
            self.optionmenu_mt.place_forget()
        self.hidden = not self.hidden
    def disable_advance(self):
        self.advanbutton['state']='disabled'
        if not self.hidden:
            self.advanbtn_text.set("Show Advanced Options")
            self.delete('showtext')
            self.fqentry.place_forget()
            self.radiobutton1.place_forget()
            self.radiobutton2.place_forget()
            self.radiobutton3.place_forget()
            self.radiobutton4.place_forget()
            self.radiobutton5.place_forget()
            self.radiobutton6.place_forget()
            self.optionmenu_mv.place_forget()
            self.optionmenu_mt.place_forget()
            self.hidden=True 
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
        if self.fqvar2.get() == '':
            self.fqvar2.set(1)
    def configure_all(self,state):
        self.cbutton['state']=state
        self.template_browse_button['state']=state
        self.template_clear_button['state']=state
        self.advanbutton['state']=state
        self.fbutton['state']=state
        self.tbutton['state']=state
        self.event_duration_menu['state']=state
        self.report_type_menu['state']=state
        self.fqentry2['state']=state
        self.navbar.optionmenu['state']=state
        self.timezonemenu['state']=state
        if not self.hidden:
            self.radiobutton1['state']=state
            self.radiobutton2['state']=state
            self.radiobutton3['state']=state
            self.radiobutton4['state']=state
            self.radiobutton5['state']=state
            self.radiobutton6['state']=state
            self.fqentry['state']=state
            self.optionmenu_mv['state']=state
            self.optionmenu_mt['state']=state
    def date_format(self,date):
        return(str(datetime.datetime.strptime(str(date),"%Y-%m-%d").strftime("%d/%m/%Y")))
if  __name__=='__main__':
    root=tk.Tk()
    root.iconbitmap(default='./img/icon.ico')
    MainApplication(root).pack(side='top',fill='both',expand=True)
    root.mainloop()

