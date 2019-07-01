"""
Application to extract data from a table on a mysql database based on a predifned frequency interval
and load this data on an excel spreadsheet.
Credits:

Icon made by VisualPharm at icon-icons.com/icon/database-the-application/2803 (License : CC Attribution)
"""
from tkcalendar import *
import xlsscript
import threading
w=800
h=500


class MainApplication(tk.Frame):
    def __init__(self,parent):
        tk.Frame.__init__(self,parent)
        self.parent=parent
        self.setup()
        self.cv=tk.Canvas(parent,height=h,width=w,bg='blue')
        self.cv.pack(side='top',fill='both',expand='yes')
        self.cv.create_image(0,0,image=self.photoimage,anchor='nw')
        self.cv.create_text(15,20,text="Time Period :",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.fbutton=ttk.Button(self.cv, text='From', command=self.fgetdate,underline=1)
        self.fbutton.place(x=150,y=20)
        self.tbutton=ttk.Button(self.cv, text='To', command=self.tgetdate)
        self.tbutton.place(x=250,y=20)
        self.tbutton.config(state='disabled')
        self.flabel=tk.Label(parent,textvariable=self.fstr,width=10)
        self.flabel.place(x=150,y=45)
        self.tlabel=tk.Label(parent,textvariable=self.tstr,width=10)
        self.tlabel.place(x=250,y=45)
        self.advanbutton=tk.Button(self.parent,textvariable=self.btn_text,command=self.advance)
        self.advanbutton.place(x=350,y=25)
        self.pentry=tk.Entry(parent)
        self.pentry.config(fg='grey')
        self.pentry.insert(0,self.default_text)
        self.pentry.bind('<FocusIn>',self.in_focus)
        self.pentry.bind('<FocusOut>',self.out_focus)
        self.cv.create_text(15,90,text="Object Path: ",fill='white',anchor='nw',font=("Arial", 12, "bold"))
        self.pentry.place(x=150,y=90,width=370)
        self.cbutton=tk.Button(parent,text="Create excel File!",command=self.extract)
        self.cbutton.place(x=200,y=240)
        self.progress = ttk.Progressbar(parent, orient=tk.HORIZONTAL,length=100,  mode='indeterminate')

    def extract(self):
        def thread_extract():
            self.progress.place(x=200,y=300)
            self.progress.start()
            self.object_fullpath=self.pentry.get()
            if self.hidden:
                found=xlsscript.main(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day),datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day),0,0,self.hidden,self.object_fullpath.upper())
            else:
                found=xlsscript.main(datetime.datetime(self.fdate.year,self.fdate.month,self.fdate.day),datetime.datetime(self.tdate.year,self.tdate.month,self.tdate.day),self.fqentry.get(),self.radb.get(),self.hidden,self.object_fullpath.upper())       
            self.progress.stop()
            self.progress.place_forget()
            self.enable_all()
            if found==-1:
                tk.messagebox.showerror("Error","Close the excel file and try again.")
            if found==-2:
                tk.messagebox.showerror("Error","Something went wrong. Restart the application and try again.")
            elif found==0:
                tk.messagebox.showwarning("Info","No data was found in the selected time period.")
            elif found:
                tk.messagebox.showinfo("Extraction Successful !","You can view the records in the file tables.xlsx")
        self.disable_all()        
        threading.Thread(target=thread_extract).start()
        
    def setup(self):
        self.tdate=calendar.datetime.date.today()
        self.fdate=calendar.datetime.date(self.tdate.year,self.tdate.month,self.tdate.day)+datetime.timedelta(days=-2)
        self.photoimage=tk.PhotoImage(file="C:\\Users\\OISM\\Desktop\\sqlApp\\bgimage.png")
        self.parent.geometry("%dx%d" % (w,h))
        self.parent.title("Create Exel Log File")
        self.btn_text=tk.StringVar(self.parent,value="Show Advanced Options")
        self.fqvar=tk.StringVar(self.parent,value='1')
        self.fqvar.trace('w',self.update)
        self.radb=tk.IntVar(self.parent,2)
        self.flag=tk.IntVar(self.parent,0)
        self.flag.trace('w',self.callback)
        self.fstr=tk.StringVar(self.parent,self.date_format(self.fdate))
        self.tstr=tk.StringVar(self.parent,self.date_format(self.tdate))
        self.default_text='MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN'
        self.hidden=True
    def fgetdate(self):
        def print_sel():
            if self.cal.selection_get()<calendar.datetime.date.today():
                self.flag.set(1)
                self.fdate=self.cal.selection_get()
                self.fstr.set(self.date_format(self.fdate))
                self.top.destroy()
            else:
                messagebox.showerror("Date Error","Date is invalid. Try Again")

        self.top = tk.Toplevel(self.parent)
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
        self.top = tk.Toplevel(self.parent)
        self.top.grab_set()
        self.cal = Calendar(self.top, font="Arial 14", selectmode='day',
                        cursor="hand1",year=self.tdate.year,month=self.tdate.month,day=self.tdate.day)
        self.cal.pack(fill="both", expand=True)
        ttk.Button(self.top, text="Go", command=print_sel).pack()
    def advance(self):
        if self.hidden:
            self.radiobuttton1=tk.Radiobutton(self.parent,text="Second(s)",value=1,variable=self.radb)
            self.radiobuttton2=tk.Radiobutton(self.parent,text="Minute(s)",value=2,variable=self.radb)
            self.radiobuttton3=tk.Radiobutton(self.parent,text="Hour(s)",value=3,variable=self.radb)
            self.fqentry=tk.Entry(self.parent,width=2,textvariable=self.fqvar)
            self.reg=self.parent.register(self.valid)
            self.fqentry.config(validate='key',validatecommand=(self.reg,'%P'))
            self.btn_text.set("Hide Advanced Options")
            self.cv.create_text(15,130,text="Show changes every ",fill='white',anchor='nw',font=("Arial",12,'bold'),tag='showtext')  
            self.fqentry.place(x=180,y=133)
            self.radiobuttton1.place(x=200,y=130)
            self.radiobuttton2.place(x=280,y=130)
            self.radiobuttton3.place(x=360,y=130)
        else:
            self.btn_text.set("Show Advanced Options")
            self.cv.delete('showtext')
            self.cv.delete('defaultinfo')
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
    def in_focus(self,*args):
        if self.pentry.get()==self.default_text:
            self.pentry.delete(0,tk.END)
            self.pentry.config(fg='black')
    def out_focus(self,*args):
        if self.pentry.get().isspace() or self.pentry.get()=='' :
            self.pentry.insert(0,self.default_text)
            self.pentry.config(fg='grey')
    def disable_all(self):
        self.cbutton['state']='disabled'
        self.advanbutton['state']='disabled'
        self.fbutton['state']='disabled'
        self.tbutton['state']='disabled'
        self.pentry['state']='disabled'
        if not self.hidden:
            self.radiobuttton1['state']='disabled'
            self.radiobuttton2['state']='disabled'
            self.radiobuttton3['state']='disabled'
            self.fqentry['state']='disabled'
    def enable_all(self):
        self.cbutton['state']='normal'
        self.advanbutton['state']='normal'
        self.fbutton['state']='normal'
        self.pentry['state']='normal'
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

