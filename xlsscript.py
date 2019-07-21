from openpyxl import Workbook,load_workbook
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.styles import Alignment,Font,PatternFill
import openpyxl.utils.cell as cell
from shutil import copyfile
import datetime
import re
from sqlscript import GetSignalData 
class Increment():
    def __init__(self,vv):
        if(isinstance(vv,float)):
            self.cur=vv
            self.inc=[0]
        else:
            self.inc=[vv]
            self.cur=None
    def next(self,vv):
        if(isinstance(vv,float)):
            if self.cur:
                self.inc.append(vv-self.cur)
                self.cur=vv
            else:
                self.inc.append(0)
                self.cur=vv                                    
        else:
            self.inc.append(vv)

class ExcessiveDataError(Exception):
    pass
class ParseData:
    def __init__(self,fdate,tdate,hidden,object_fullpaths,signal_type,fq=None,radb=None,data_type='Changes',template_path=None):
        try:
            self.str_time=datetime.datetime.now().strftime('%a %d %b %y %I-%M-%S%p')
            self.file_path = './SignalLog/'+self.str_time+'.xlsx'
            self.template_path = template_path
            self.signal_info=[]
            self.all_signal_data=[]
            self.object_fullpaths = list(object_fullpaths)
            for path in self.object_fullpaths:
                self.signal_info.append(GetSignalData(path,signal_type))
            for signal in self.signal_info:
                self.data=[]
                self.unparsed_data = signal.result
                self.uid = signal.uid
                if not self.unparsed_data:
                    self.all_signal_data.append([])
                    continue
                elif signal_type in ['All Measurements','All Metering']:
                    if not hidden:
                        self.option={1:1.0,2:60.0,3:3600.0}
                        self.roundTo=(float)(fq)*self.option.get(radb)
                        self.found_first={self.uid:0}
                        next_dt = fdate
                        cur_val=()
                        for vm,vo,dt,ms,vq,vv in self.unparsed_data:
                            dt=dt+datetime.timedelta(milliseconds=ms)
                            if len(self.data)>12800:
                                raise ExcessiveDataError
                            if vo != self.uid:
                                continue
                            elif dt>tdate:
                                break
                            elif dt>next_dt:
                                if self.found_first[vo] == 0:
                                    if cur_val:
                                        self.incrementdata = Increment(cur_val[3])
                                        self.average_aux = [float(cur_val[3]),1.0,float(cur_val[3]),float(cur_val[3])]
                                        self.avg = [(self.average_aux[0]/self.average_aux[1],float(cur_val[3]),float(cur_val[3])),]
                                        self.data.append((str(next_dt.date()),str(next_dt.time()),*cur_val),)
                                    else:
                                        cur_val=('Disconnected',)*4
                                        self.incrementdata=Increment('Disconnected')
                                        self.average_aux=[None,None,None,None]
                                        self.avg=[('Disconnected',)*3]
                                        self.data.append((str(next_dt.date()),str(next_dt.time()),*cur_val),)
                                    self.found_first[vo]=1
                                else:
                                        self.incrementdata.next(cur_val[3])
                                        self.avg.append((self.average_aux[0]/self.average_aux[1],self.average_aux[2],self.average_aux[3]))
                                        self.data.append((str(next_dt.date()),str(next_dt.time()),*cur_val),)
                                next_dt=self.roundTime(next_dt,self.roundTo)                            
                                while dt>next_dt and next_dt<tdate:
                                    self.incrementdata.next('Disconnected')
                                    self.avg.append(('Disconnected',)*3)
                                    self.data.append((str(next_dt.date()),str(next_dt.time()),*cur_val),)
                                    next_dt=self.roundTime(next_dt,self.roundTo)
                                cur_val=(vm,vo,vq,vv)
                                self.average_aux[0]=self.average_aux[2]=self.average_aux[3]=vv
                                self.average_aux[1]=1
                            elif dt<=next_dt:
                                cur_val=(vm,vo,vq,vv)
                                if self.found_first[vo] == 0:
                                    continue
                                else:
                                    if vv>self.average_aux[2]:
                                        self.average_aux[2]=vv
                                    elif vv<self.average_aux[3]:
                                        self.average_aux[3]=vv
                                    self.average_aux[0]+=vv
                                    self.average_aux[1]+=1
                        if self.data:
                            if data_type == 'Average':
                                self.data = [i+j for i,j in zip(self.data,self.avg)]
                            elif data_type == 'Consumption':
                                self.data = [i+(k,) for i,k in zip(self.data,self.incrementdata.inc)]
                    else:
                        for vm,vo,dt,ms,vq,vv in self.unparsed_data:
                            if(dt>=fdate and dt<=tdate and  vo==self.uid):
                                self.data.append((str(dt.date()),str(self.addSecs(dt.time(),ms)),vm,vo,vq,vv))
                else:
                    for dt,ms,eo,em,eu in self.unparsed_data:
                        if(dt>=fdate and dt<=tdate and  eo==self.uid):
                            match=re.findall('.+=(.+)',eu)
                            if match:
                                match=match[0]
                            else:
                                match=''
                            self.data.append((str(dt.date()),str(self.addSecs(dt.time(),ms)),em,match))
                self.all_signal_data.append(self.data)
            if not self.isListEmpty(self.all_signal_data):
                self.create_table(signal_type,data_type)
                self.populate_table()
                if self.not_found:
                    self.result = 2
                else:
                    self.result = 1
            else:
                self.result = 0
        except PermissionError:
            self.result = -1
        except ExcessiveDataError:
            self.result = -3
        except Exception as e:
            self.result = -2
            print(e)
    def populate_table(self):
        pass

    def create_table(self,signal_type,data_type):
        def header_build(step,column=3):
            for i,path in enumerate(self.object_fullpaths,start=1):
                for j in range(step):
                    self.ws.column_dimensions[cell.get_column_letter(column+j)].width=25
                base_cell = self.ws.cell(row=self.cur_row, column=column)
                self.ws.merge_cells(start_row=self.cur_row,end_row=self.cur_row+2,start_column=column,end_column=column+step-1)
                base_cell.value = path
                base_cell.alignment = align_head
                if i%2:            
                    base_cell.font = alt_font_head
                    base_cell.fill = alt_fill_head
                else:
                    base_cell.font = font_head
                    base_cell.fill = fill_head
                self.ws.column_dimensions.group(cell.get_column_letter(column),cell.get_column_letter(column+step-1),hidden=False,outline_level=i)
                column+=step
            self.ws.cell(row=self.cur_row+3,column=1,value='Date')
            self.ws.cell(row=self.cur_row+3,column=2,value='Time')
            self.ws.column_dimensions['A'].width=15
            self.ws.column_dimensions['B'].width=15
        def type1(step=4):
            header_build(step)
            column = 3
            for i,field in enumerate([['Value Meantype-','Value Object UID-','Value Quality-','Value-']]*len(self.object_fullpaths),start=1):
                for j in range(step):
                    self.ws.cell(row=self.cur_row+3,column=column+j,value=field[j]+str(i))
                column+=step
            return step
        def type2(step=5):
            header_build(step)
            column = 3
            for i,field in enumerate([['Value Meantype-','Value Object UID-','Value Quality-','Value-','Value Consumption-']]*len(self.object_fullpaths),start=1):
                for j in range(step):
                    self.ws.cell(row=self.cur_row+3,column=column+j,value=field[j]+str(i))
                column+=step
            return step
        def type3(step=7):
            header_build(step)
            column = 3
            for i,field in enumerate([['Value Meantype-','Value Object UID-','Value Quality-','Value-','Value Average-','Value Max-','Value Min-']]*len(self.object_fullpaths),start=1):
                for j in range(step):
                    self.ws.cell(row=self.cur_row+3,column=column+j,value=field[j]+str(i))
                column+=step
            return step
        def type4(step=2):
            self.ws.append(['Date','Time','Event Mess','Event Userlooked'],start=1)
            header_build(step)
            column = 3
            for i,field in enumerate([['Value Meantype-','Value Object UID-','Value Quality-','Value Value-']]*len(self.object_fullpaths)):
                for j in range(step):
                    self.ws.cell(row=self.cur_row+3,column=column+j,value=field[j]+str(i))
                column+=step
            return step
        if self.template_path:
            copyfile(self.template_path,self.file_path)
            self.wb = load_workbook(filename=self.file_path)
        else:
            self.wb = Workbook()
        self.ws = self.wb.active
        self.cur_row = self.ws.max_row+2
        font_head =  Font(b=True,color="4EB1BA",sz=15)
        alt_font_head = Font(b=True,color='FFFFFF',sz=15)
        align_head = Alignment(horizontal='center',vertical='center')
        fill_head = PatternFill('solid',fgColor='222930')
        alt_fill_head = PatternFill('solid',fgColor='FF8362')
        self.zipped_data = []
        self.not_found=[]
        table_type = {'All Measurements':{'Changes':type1,'Consumption':type2,'Average':type3}}

        for index,signal_data in enumerate(self.all_signal_data):
            if signal_data and self.zipped_data == []:
                self.zipped_data = signal_data
            elif signal_data and self.zipped_data:
                self.zipped_data = [i+j[2:] for i,j in zip(self.zipped_data,signal_data)]
            elif not signal_data:
                self.not_found.append(self.object_fullpaths.pop(index))
        step = table_type.get(signal_type,dict()).get(data_type,type4)()

        for row in self.zipped_data:
            self.ws.append(row)

        tab = Table(displayName='Report',ref="A"+str(self.cur_row+3)+":"+cell.get_column_letter(2+len(self.object_fullpaths)*step)+str(len(self.zipped_data)+4))
        style = TableStyleInfo(name='TableStyleMedium9',showColumnStripes=True,showRowStripes=False)
        tab.tableStyleInfo = style
        self.ws.add_table(tab)
        self.wb.save(self.file_path)
        self.wb.close()
    def roundTime(self,d1,roundTo):
        seconds=(d1-d1.min).seconds
        rounding=(seconds+roundTo/2)//roundTo*roundTo
        d2=d1+datetime.timedelta(0,rounding-seconds,-d1.microsecond)
        if(d2>d1):
            return d2
        else:
            return d2+datetime.timedelta(seconds=roundTo)
    def addSecs(self,tm,ms):
        d1=datetime.datetime(1,1,1,tm.hour,tm.minute,tm.second)
        d1+=datetime.timedelta(milliseconds=ms)
        return d1.time()
    def isListEmpty(self,inList):
        if isinstance(inList,list):
            return all(map(self.isListEmpty,inList))
        return False



if __name__ == "__main__":
  #a=ParseData(datetime.datetime(2019,5,28),datetime.datetime(2019,5,30),2,1,False,'MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN','All Signals')
  #b=ParseData(datetime.datetime(2019,5,29,9,0),datetime.datetime(2019,6,10,9,1),False,['MOSG / 33KV / H03_CABLEFDR-H16 / MEASUREMENT / VOLTAGE VBR','MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN'],'All Measurements',1,1,'Average')
  #b=ParseData(datetime.datetime(2019,5,28),datetime.datetime(2019,7,6),1,1,True,'MOSG / 33KV / H38_39 BUS SEC / GIS SIGNALS / GIS VT  MCB TRIP','All Signals')
  b=ParseData(datetime.datetime(2019,7,9,14,5),datetime.datetime(2019,7,9,14,14),False,[
    'MOSG / 33KV / H38_39 BUS SEC / MEASUREMENT / ACTIVE POWER(P)',
    'MOSG / 33KV / H38_39 BUS SEC / MEASUREMENT / FREQUENCY',
    'MOSG / 33KV / H38_39 BUS SEC / MEASUREMENT / CURRENT IB',
    ],'All Measurements',5,1,'Average') #19,25,4
  print(b.result)
 