from openpyxl import Workbook,load_workbook
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.styles import Alignment,Font,PatternFill
import openpyxl.utils.cell as cell
from shutil import copyfile
import datetime
import re
from os import path as os_path
from os import mkdir
from sqlscript import GetSignalData 
import pickle
options_default = {
    'Excel':{
        1:0
    },
    'Template':{
        1:''
    }
}
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
    def __init__(self,fdate,tdate,timezone,hidden,object_fullpaths,signal_type,fq=None,radb=None,data_type='Changes',template_path=None):
        try:
            self.str_time=datetime.datetime.now().strftime('%a %d %b %y %I-%M-%S%p')
            if not os_path.isdir('./SignalLog'):
                mkdir('./SignalLog')
            self.file_path = './SignalLog/'+self.str_time+'.xlsx'
            self.template_path = template_path #= "C:/Users/OISM/Desktop/Signal Report/Templates/template1.xlsx"
            timezonediff = (-1)*int(re.findall('GMT(.+):00',timezone)[0])
            fdate += datetime.timedelta(hours=timezonediff)
            tdate += datetime.timedelta(hours=timezonediff)
            timezonediff*=-1
            if os_path.isfile('settings'):
                with open('settings','rb') as file:
                    self.customization = pickle.load(file)
            else:
                self.customization = options_default
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
                            dt+=datetime.timedelta(milliseconds=ms)
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
                                        self.average_aux = [
                                            float(cur_val[3]),
                                            1.0,
                                            float(cur_val[3]),
                                            float(cur_val[3])
                                        ]
                                        self.avg = [
                                            (
                                                self.average_aux[0]/self.average_aux[1],
                                                float(cur_val[3]),
                                                float(cur_val[3])
                                            ),
                                        ]
                                        self.data.append(
                                            (
                                                str((next_dt+datetime.timedelta(hours=timezonediff)).date()),
                                                str((next_dt+datetime.timedelta(hours=timezonediff)).time()),
                                                *cur_val,
                                            ),
                                        )
                                    else:
                                        cur_val=('Disconnected',)*4
                                        self.incrementdata=Increment('Disconnected')
                                        self.average_aux=[None,None,None,None]
                                        self.avg=[('Disconnected',)*3]
                                        self.data.append(
                                            (
                                                str((next_dt+datetime.timedelta(hours=timezonediff)).date()),
                                                str((next_dt+datetime.timedelta(hours=timezonediff)).time()),
                                                *cur_val
                                            ),
                                        )
                                    self.found_first[vo]=1
                                else:
                                        self.incrementdata.next(cur_val[3])
                                        self.avg.append(
                                            (
                                                self.average_aux[0]/self.average_aux[1],
                                                self.average_aux[2],
                                                self.average_aux[3],
                                            ),
                                        )
                                        self.data.append(
                                            (
                                                str((next_dt+datetime.timedelta(hours=timezonediff)).date()),
                                                str((next_dt+datetime.timedelta(hours=timezonediff)).time()),
                                                *cur_val                                               
                                            ),
                                        )
                                next_dt=self.roundTime(next_dt,self.roundTo)                            
                                while next_dt<dt and next_dt<tdate:
                                    self.incrementdata.next('Disconnected')
                                    self.avg.append(('Disconnected',)*3)
                                    self.data.append(
                                        (
                                            str((next_dt+datetime.timedelta(hours=timezonediff)).date()),
                                            str((next_dt+datetime.timedelta(hours=timezonediff)).time()),
                                            *cur_val                                               
                                        ),
                                    )
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
                        while next_dt<=tdate:
                            self.incrementdata.next('Disconnected')
                            self.avg.append(('Disconnected',)*3)
                            self.data.append(
                                (
                                    str((next_dt+datetime.timedelta(hours=timezonediff)).date()),
                                    str((next_dt+datetime.timedelta(hours=timezonediff)).time()),
                                    *cur_val                                               
                                ),
                            )
                            next_dt=self.roundTime(next_dt,self.roundTo)

                        if self.data:
                            if data_type == 'Average':
                                self.data = [i+j for i,j in zip(self.data,self.avg)]
                            elif data_type == 'Consumption':
                                self.data = [i+(k,) for i,k in zip(self.data,self.incrementdata.inc)]
                    else:
                        for vm,vo,dt,ms,vq,vv in self.unparsed_data:
                            dt+=datetime.timedelta(milliseconds=ms)
                            if(dt>=fdate and dt<=tdate and  vo==self.uid):
                                self.data.append(
                                    (
                                        str((dt+datetime.timedelta(hours=timezonediff)).date()),
                                        str((dt+datetime.timedelta(hours=timezonediff)).time()),
                                        vm,vo,vq,vv
                                    ),
                                )
                else:
                    for dt,ms,eo,em,eu in self.unparsed_data:
                        dt+=datetime.timedelta(milliseconds=ms)
                        if(dt>=fdate and dt<=tdate and  eo==self.uid):
                            match=re.findall('.+=(.+)',eu)
                            if match:
                                match=match[0]
                            else:
                                match=''
                            self.data.append(
                                (
                                    str((dt+datetime.timedelta(hours=timezonediff)).date()),
                                    str((dt+datetime.timedelta(hours=timezonediff)).time()),
                                    em,match
                                )
                            )
                self.all_signal_data.append(self.data)
            if not self.isListEmpty(self.all_signal_data):
                self.create_table(signal_type,data_type,hidden)
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
            self.errormessage = str(e)
            print(e)


    def create_table(self,signal_type,data_type,hidden):
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
                if not self.customization['Excel'][1]:
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
        def type4(index,step):
            column = 1
            for j in range(1,step+1):
                self.ws.column_dimensions[cell.get_column_letter(j)].width=25
            base_cell = self.ws.cell(row=self.cur_row, column=column)
            self.ws.merge_cells(start_row=self.cur_row,end_row=self.cur_row+2,start_column=column,end_column=column+step-1)
            base_cell.value = self.object_fullpaths[index]
            base_cell.alignment = align_head
            if index%2:            
                base_cell.font = alt_font_head
                base_cell.fill = alt_fill_head
            else:
                base_cell.font = font_head
                base_cell.fill = fill_head
            if step==6:
                for i,head in enumerate(['Date','Time','Value Meantype','Value Object UID','Value Quality','Value']):
                    self.ws.cell(row=self.cur_row+3,column=column+i,value=head)
            else:
                for i,head in enumerate(['Date','Time','Event Mess','Event Userlooked']):
                    self.ws.cell(row=self.cur_row+3,column=column+i,value=head)
            for row in self.all_signal_data[index]:
                self.ws.append(row)
            tab = Table(displayName='Report'+str(index),ref="A"+str(self.cur_row+3)+":"+cell.get_column_letter(step)+str(len(self.all_signal_data[index])+self.cur_row+3))
            style = TableStyleInfo(name='TableStyleMedium9',showColumnStripes=True,showRowStripes=False)
            tab.tableStyleInfo = style
            self.ws.add_table(tab)
            self.ws = self.wb.create_sheet()
            self.cur_row = self.ws.max_row 

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
        self.not_found=[]
        if not hidden:
            self.zipped_data = []
            table_type = dict.fromkeys(
                ['All Measurements','All Metering'], {
                    'Changes':type1,
                    'Consumption':type2,
                    'Average':type3,
                }
            )
            for index,signal_data in enumerate(self.all_signal_data):
                if signal_data and self.zipped_data == []:
                    self.zipped_data = signal_data
                elif signal_data and self.zipped_data:
                    self.zipped_data = [i+j[2:] for i,j in zip(self.zipped_data,signal_data)]
                elif not signal_data:
                    self.not_found.append(self.object_fullpaths.pop(index))
            step = table_type.get(signal_type,dict()).get(data_type,None)()

            for row in self.zipped_data:
                self.ws.append(row)

            tab = Table(displayName='Report',ref="A"+str(self.cur_row+3)+":"+cell.get_column_letter(2+len(self.object_fullpaths)*step)+str(len(self.zipped_data)+self.cur_row+3))
            style = TableStyleInfo(name='TableStyleMedium9',showColumnStripes=True,showRowStripes=False)
            tab.tableStyleInfo = style
            self.ws.add_table(tab)       
        else:
            for index,signal_data in enumerate(self.all_signal_data):
                if signal_data:
                    if signal_type in ['All Measurements','All Metering']:
                        type4(index,step=6)
                    else:
                        type4(index,step=4)
                else:
                    self.not_found.append(self.object_fullpaths[index])
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
    def isListEmpty(self,inList):
        if isinstance(inList,list):
            return all(map(self.isListEmpty,inList))
        return False



if __name__ == "__main__":

    b=ParseData(datetime.datetime(2019,7,22,12,39),datetime.datetime(2019,7,22,12,55),'GMT+4:00',False,[
      'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / ACTIVE POWER(P)',
      'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / CURRENT IB',
      'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / VOLTAGE VRY',
      'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / FREQUENCY'
  ],'All Measurements',5,1,'Average') #65,2,25 
    # c=ParseData(datetime.datetime(2019,7,4,10,41),datetime.datetime(2019,7,4,11,0),'GMT+4:00',True,[
    #     'MOSG / 33KV / H12_13 BUS SEC / Q0 CB / POSITION'
    # ],'All Signals')
    print(b.result)
 