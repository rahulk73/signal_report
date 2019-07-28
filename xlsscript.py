from openpyxl import Workbook,load_workbook
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.styles import Alignment,Font,PatternFill
import openpyxl.utils.cell as cell
import datetime
import re
from os import path as os_path
from os import mkdir
from sqlscript import GetSignalData 
import pickle
class Preferences:
    options_default = {
    'Excel':{
        1:0,
        2:0,
        3:4,
    },
    'Template':{
        1:'',
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
                self.customization = Preferences.options_default
            self.chart_counter = 0
            self.changes_mv_fields = ('Value Meantype','Value Object UID','Value Quality','Measurement')
            self.consumption_mt_fields = ('Value Meantype','Value Object UID','Value Quality','Metering','Consumption')
            self.average_mv_fields = ('Value Meantype','Value Object UID','Value Quality','Measurement','Average','Max','Min')
            self.changes_event_fields = ('Event Mess','Event Userlooked')
            self.font_head =  Font(b=True,color="4EB1BA",sz=15)
            self.alt_font_head = Font(b=True,color='FFFFFF',sz=15)
            self.align_head = Alignment(horizontal='center',vertical='center')
            self.fill_head = PatternFill('solid',fgColor='222930')
            self.alt_fill_head = PatternFill('solid',fgColor='FF8362')
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
                        else:
                            if self.found_first[self.uid] == 0:
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
                                self.found_first[self.uid]=1
                                next_dt=self.roundTime(next_dt,self.roundTo)
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
                if not hidden and not self.customization['Excel'][2]:
                    self.create_one_sheet_table(signal_type,data_type)
                else:
                    self.create_multiple_sheet_table(signal_type,data_type,hidden)
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
    def create_chart(self,column_start,row_start,row_end,column_end=None):
        def type1():
            from openpyxl.chart import AreaChart
            chart = AreaChart()
            return chart
        def type2():
            from openpyxl.chart import BarChart
            chart = BarChart()
            chart.type='col'
            return chart
        def type3():
            from openpyxl.chart import BarChart
            chart = BarChart()
            chart.type='bar'
            return chart
        def type4():
            from openpyxl.chart import LineChart
            # from openpyxl.chart.axis import DateAxis
            chart = LineChart()
            return chart
        def type0():
            return
        self.chart_options = {
            1:type1,
            2:type2,
            3:type3,
            4:type4,
            0:type0,
        }
        chart = self.chart_options[self.customization['Excel'][3]]()
        from openpyxl.chart import Reference
        chart.style = 42
        chart.title = self.object_fullpaths[self.chart_counter]
        self.chart_counter+=1
        data = Reference(self.ws,min_col=column_start,max_col=column_end,min_row=row_start,max_row=row_end)
        chart.add_data(data,titles_from_data=True)
        cats = Reference(self.ws,min_col=1,max_col=2,min_row=row_start+1,max_row=row_end)
        chart.set_categories(cats)
        self.ws.add_chart(chart,cell.get_column_letter(column_start)+str(self.ws.max_row+1))
    def create_one_sheet_table(self,signal_type,data_type):
        def header_build(step,column=3):
            for i,path in enumerate(self.object_fullpaths,start=1):
                for j in range(step):
                    self.ws.column_dimensions[cell.get_column_letter(column+j)].width=25
                base_cell = self.ws.cell(row=self.cur_row, column=column)
                self.ws.merge_cells(start_row=self.cur_row,end_row=self.cur_row+2,start_column=column,end_column=column+step-1)
                base_cell.value = path
                base_cell.alignment = self.align_head
                if i%2:            
                    base_cell.font = self.alt_font_head
                    base_cell.fill = self.alt_fill_head
                else:
                    base_cell.font = self.font_head
                    base_cell.fill = self.fill_head
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
            for i,field in enumerate([[*self.changes_mv_fields]]*len(self.object_fullpaths),start=1):
                for j in range(step):
                    self.ws.cell(row=self.cur_row+3,column=column+j,value=field[j]+str(i))
                column+=step
            return step
        def type2(step=5):
            header_build(step)
            column = 3
            for i,field in enumerate([[*self.consumption_mt_fields]]*len(self.object_fullpaths),start=1):
                for j in range(step):
                    self.ws.cell(row=self.cur_row+3,column=column+j,value=field[j]+str(i))
                column+=step
            return step
        def type3(step=7):
            header_build(step)
            column = 3
            for i,field in enumerate([[*self.average_mv_fields]]*len(self.object_fullpaths),start=1):
                for j in range(step):
                    self.ws.cell(row=self.cur_row+3,column=column+j,value=field[j]+'-'+str(i))
                column+=step
            return step
        if self.template_path:
            self.wb = load_workbook(filename=self.template_path)
            self.ws_template = self.wb.active
            self.ws = self.wb.copy_worksheet(self.ws_template)
            self.wb.remove(self.ws_template)
            self.ws.title = 'Report'
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
        self.cur_row = self.ws.max_row+2
        self.not_found=[]
        self.zipped_data = []
        table_type = {
            'All Measurements': {
                'Changes':type1,
                'Average':type3,
            },
            'All Metering': {
                'Changes':type1,
                'Consumption':type2
            },
        }
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
        
        table_row_start = self.cur_row+3
        table_column_end = 2+len(self.object_fullpaths)*step
        table_row_end = len(self.zipped_data)+self.cur_row+3

        tab = Table(displayName='Report',ref="A"+str(table_row_start)+":"+cell.get_column_letter(table_column_end)+str(table_row_end))
        style = TableStyleInfo(name='TableStyleMedium9',showColumnStripes=True,showRowStripes=False)
        tab.tableStyleInfo = style
        self.ws.add_table(tab)
        if self.customization['Excel'][3] and signal_type in ['All Measurements','All Metering']:
            for i in range(6,6+step*(len(self.object_fullpaths)-1)+1,step):
                if data_type in ['Average','Consumption']:
                    self.create_chart(column_start=i,column_end=i+1,row_start=table_row_start,row_end=table_row_end)
                else:
                    self.create_chart(column_start=i,row_start=table_row_start,row_end=table_row_end)
        self.wb.save(self.file_path)
        self.wb.close()

    def create_multiple_sheet_table(self,signal_type,data_type,hidden):
        self.sheet_index = 1
        def build(index,step):
            column = 1
            for j in range(1,step+1):
                self.ws.column_dimensions[cell.get_column_letter(j)].width=25
            base_cell = self.ws.cell(row=self.cur_row, column=column)
            self.ws.merge_cells(start_row=self.cur_row,end_row=self.cur_row+2,start_column=column,end_column=column+step-1)
            base_cell.value = self.object_fullpaths[index]
            base_cell.alignment = self.align_head
            if index%2:            
                base_cell.font = self.alt_font_head
                base_cell.fill = self.alt_fill_head
            else:
                base_cell.font = self.font_head
                base_cell.fill = self.fill_head
            if step==6:
                for i,head in enumerate(['Date','Time',*self.changes_mv_fields]):
                    self.ws.cell(row=self.cur_row+3,column=column+i,value=head)
            elif step==7:
                for i,head in enumerate(['Date','Time',*self.consumption_mt_fields]):
                    self.ws.cell(row=self.cur_row+3,column=column+i,value=head)
            elif step==9:
                for i,head in enumerate(['Date','Time',*self.average_mv_fields]):
                    self.ws.cell(row=self.cur_row+3,column=column+i,value=head)         
            else:
                for i,head in enumerate(['Date','Time',*self.changes_event_fields]):
                    self.ws.cell(row=self.cur_row+3,column=column+i,value=head)
            for row in self.all_signal_data[index]:
                self.ws.append(row)

            table_row_start = self.cur_row+3
            table_column_end = step
            table_row_end = len(self.all_signal_data[index])+self.cur_row+3

            tab = Table(displayName='Report'+str(index),ref="A"+str(table_row_start)+":"+cell.get_column_letter(table_column_end)+str(table_row_end))
            style = TableStyleInfo(name='TableStyleMedium9',showColumnStripes=True,showRowStripes=False)
            tab.tableStyleInfo = style
            self.ws.add_table(tab)
            if signal_type in ['All Measurements','All Metering']:
                if data_type in ['Average','Consumption']:
                    self.create_chart(column_start=6,column_end=7,row_start=table_row_start,row_end=table_row_end)
                else:
                    self.create_chart(column_start=step,row_start=table_row_start,row_end=table_row_end)
            if self.template_path:
                self.ws = self.wb.copy_worksheet(self.ws_template)
                self.ws.title = str(self.sheet_index)
                self.sheet_index+=1
            else:
                self.ws = self.wb.create_sheet()
            self.cur_row = self.ws.max_row+2 

        if self.template_path:
            self.wb = load_workbook(filename=self.template_path)
            self.ws_template = self.wb.active
            self.ws = self.wb.copy_worksheet(self.ws_template)
            self.wb.remove(self.ws_template)
            self.ws.title = str(self.sheet_index)
            self.sheet_index+=1
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
        self.cur_row = self.ws.max_row+2
        self.not_found=[]
        table_type = {
            'All Measurements': {
                'Changes':6,
                'Average':9,
            },
            'All Metering': {
                'Changes':6,
                'Consumption':7,
            },
        }
        for index,signal_data in enumerate(self.all_signal_data):
            if signal_data:
                build(index,step=table_type.get(signal_type,dict()).get(data_type,4))
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

#     b=ParseData(datetime.datetime(2019,7,22,12,54),datetime.datetime(2019,7,22,12,55),'GMT+4:00',False,[
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / ACTIVE POWER(P)',
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / CURRENT IB',
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / VOLTAGE VRY',
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / FREQUENCY'
#   ],'All Measurements',5,1,'Average') #65,2,25 
#     b=ParseData(datetime.datetime(2019,7,22,12,39),datetime.datetime(2019,7,22,12,55),'GMT+4:00',True,[
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / ACTIVE POWER(P)',
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / CURRENT IB',
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / VOLTAGE VRY',
#       'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / FREQUENCY'
#   ],'All Measurements') #65,2,25 
    c=ParseData(datetime.datetime(2019,7,4,10,41),datetime.datetime(2019,7,4,11,0),'GMT+4:00',True,[
        'MOSG / 33KV / H12_13 BUS SEC / Q0 CB / POSITION'
    ],'All Signals')
    print(c.result)
 