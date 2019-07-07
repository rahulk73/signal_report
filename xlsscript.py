import xlsxwriter
import datetime
from sqlscript import GetSignalData
class ExcessiveDataError(Exception):
    pass
class ParseData:
    def __init__(self,fdate,tdate,fq,radb,hidden,object_fullpath,signal_type):
        try:    
            self.workbook = xlsxwriter.Workbook('./SignalLog/tables.xlsx')
            self.worksheet = self.workbook.add_worksheet()
            cell_format=self.workbook.add_format({'font_color':'white','bg_color':'#107896',\
        'font_name':'Times New Roman','bold':True,'underline':True,'center_across':True,'font_size':20})  
            caption = 'LOGS'
            self.worksheet.set_column('B:G', 20)
            self.worksheet.write('D1', caption,cell_format)
            self.signaldata=GetSignalData(object_fullpath,signal_type)
            self.unparsed_data=self.signaldata.result
            self.uid=self.signaldata.uid
            if not self.unparsed_data:
                self.workbook.close()
                self.result= -3
            else:
                self.data=[]
                if not hidden:
                    self.option={1:1.0,2:60.0,3:3600.0}
                    self.roundTo=(float)(fq)*self.option.get(radb)
                    self.flag={}
                    if self.unparsed_data:
                        self.flag[self.uid]=0
                    for vm,vo,dt,ms,vq,vv in self.unparsed_data:
                        dt=dt+datetime.timedelta(milliseconds=ms)
                        if len(self.data)>12800:
                            raise ExcessiveDataError
                        if dt<fdate or vo !=self.uid:
                            continue
                        if dt>tdate:
                            if self.flag[vo]==1:
                                self.flag[vo]=0
                            continue
                        elif dt >=fdate and dt < (tdate+datetime.timedelta(1)) and self.flag[vo]==0:
                            self.flag[vo]=1
                            cur_val=(vm,vo,vq,vv)
                            next_dt=self.roundTime(dt,self.roundTo)
                            self.data.append((str(dt.date()),str(self.addSecs(dt.time(),ms)),vm,vo,vq,vv))
                            continue
                        elif self.flag[vo]==1 and dt<=next_dt:
                            cur_val=(vm,vo,vq,vv)
                        elif self.flag[vo]==1 and dt>next_dt:
                            while dt>next_dt and next_dt<tdate:
                                self.data.append((str(next_dt.date()),str(next_dt.time()),*cur_val))
                                next_dt=self.roundTime(next_dt,self.roundTo)
                            cur_val=(vm,vo,vq,vv)
                else:
                    for vm,vo,dt,ms,vq,vv in self.unparsed_data:
                        if(dt>=fdate and dt<(tdate+datetime.timedelta(1)) and  vo==self.uid):
                            self.data.append((str(dt.date()),str(self.addSecs(dt.time(),ms)),vm,vo,vq,vv))
                if self.data:
                    options = {'banded_rows': 0, 'banded_columns': 1,'data': self.data,
                        'columns': [{'header':'Date'},
                                    {'header': 'Time'},
                                    {'header': 'Value Meantype'},
                                    {'header': 'Value Object UID '},
                                    {'header': 'Value Quality'},
                                    {'header': 'Value Value'},
                                    ]} 
                    self.worksheet.add_table('B3:G'+str(3+len(self.data)),options)
                self.workbook.close()
                self.result = int(not self.data==[])
        except PermissionError:
            self.result = -1
        except ExcessiveDataError:
            self.result = -4
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



if __name__ == "__main__":
  #a=ParseData(datetime.datetime(2019,5,28),datetime.datetime(2019,5,30),2,1,False,'MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN','All Signals')
  b=a=ParseData(datetime.datetime(2019,5,28),datetime.datetime(2019,7,6),1,1,False,'MOSG / 33KV / H03_CABLEFDR-H16 / MEASUREMENT / VOLTAGE VBR','All Signals')
  print(b.result)
 