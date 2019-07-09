import xlsxwriter
import datetime
import re
from sqlscript import GetSignalData 
class Increment():
    def __init__(self,vv):
        self.cur=vv
        self.inc=[0]
    def next(self,vv):
        if(isinstance(vv,float)):
            self.inc.append(vv-self.cur)
            self.cur=vv
        else:
            self.inc.append(vv)

class ExcessiveDataError(Exception):
    pass
class ParseData:
    def __init__(self,fdate,tdate,hidden,object_fullpath,signal_type,fq=None,radb=None,data_type='Changes'):
        try:    
            self.str_time=datetime.datetime.now().strftime('%a %d %b %y %I-%M-%S%p')
            self.workbook = xlsxwriter.Workbook('./SignalLog/'+self.str_time+'.xlsx')
            self.worksheet = self.workbook.add_worksheet()
            caption = object_fullpath
            self.worksheet.set_column('B:J', 20)
            self.worksheet.write('D1', caption)
            self.signaldata=GetSignalData(object_fullpath,signal_type)
            self.unparsed_data=self.signaldata.result
            self.uid=self.signaldata.uid
            if not self.unparsed_data:
                self.workbook.close()
                self.result= -3
            elif self.signaldata.mvsignal:
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
                            self.incrementdata=Increment(vv)
                            self.averagedata=[float(vv),1.0,float(vv),float(vv)]
                            self.avg=[(self.averagedata[0]/self.averagedata[1],float(vv),float(vv))]
                            next_dt=self.roundTime(dt,self.roundTo)
                            self.data.append((str(dt.date()),str(self.addSecs(dt.time(),ms)),vm,vo,vq,vv))
                            continue
                        elif self.flag[vo]==1 and dt<=next_dt:
                            cur_val=(vm,vo,vq,vv)
                            if vv>self.averagedata[2]:
                                self.averagedata[2]=vv
                            elif vv<self.averagedata[3]:
                                self.averagedata[3]=vv
                            self.averagedata[0]+=vv
                            self.averagedata[1]+=1
                        elif self.flag[vo]==1 and dt>next_dt:
                            self.incrementdata.next(cur_val[3])
                            self.avg.append((self.averagedata[0]/self.averagedata[1],self.averagedata[2],self.averagedata[3]))
                            self.data.append((str(next_dt.date()),str(next_dt.time()),*cur_val))
                            next_dt=self.roundTime(next_dt,self.roundTo)                            
                            while dt>next_dt and next_dt<tdate:
                                self.incrementdata.next('Disconnected')
                                self.avg.append(('Disconnected',))
                                self.data.append((str(next_dt.date()),str(next_dt.time()),*cur_val))
                                next_dt=self.roundTime(next_dt,self.roundTo)
                            cur_val=(vm,vo,vq,vv)
                            self.averagedata[0]=self.averagedata[2]=self.averagedata[3]=vv
                            self.averagedata[1]=1
                else:
                    for vm,vo,dt,ms,vq,vv in self.unparsed_data:
                        if(dt>=fdate and dt<(tdate+datetime.timedelta(1)) and  vo==self.uid):
                            self.data.append((str(dt.date()),str(self.addSecs(dt.time(),ms)),vm,vo,vq,vv))
                if self.data:
                    if data_type == 'Changes':
                        options = {'banded_rows': 0, 'banded_columns': 1,'data': self.data,
                            'columns': [{'header':'Date'},
                                        {'header': 'Time'},
                                        {'header': 'Value Meantype'},
                                        {'header': 'Value Object UID '},
                                        {'header': 'Value Quality'},
                                        {'header': 'Value Value'},
                                        ]} 
                        self.worksheet.add_table('B3:G'+str(3+len(self.data)),options)
                    elif data_type == 'Average':
                        zipped=[i[:-1]+j for i,j in zip(self.data,self.avg)]
                        options = {'banded_rows': 0, 'banded_columns': 1,'data': zipped,
                            'columns': [{'header':'Date'},
                                        {'header': 'Time'},
                                        {'header': 'Value Meantype'},
                                        {'header': 'Value Object UID '},
                                        {'header': 'Value Quality'},
                                        {'header': 'Value Average'},
                                        {'header': 'Value Max'},
                                        {'header': 'Value Min'},
                                        ]} 
                        self.worksheet.add_table('B3:I'+str(3+len(self.data)),options)
                    elif data_type == 'Consumption':
                        zipped=[i+(j,) for i,j in zip(self.data,self.incrementdata.inc)]
                        options = {'banded_rows': 0, 'banded_columns': 1,'data': zipped,
                            'columns': [{'header':'Date'},
                                        {'header': 'Time'},
                                        {'header': 'Value Meantype'},
                                        {'header': 'Value Object UID '},
                                        {'header': 'Value Quality'},
                                        {'header': 'Value Value'},
                                        {'header': 'Value Consumption'},
                                        ]} 
                        self.worksheet.add_table('B3:H'+str(3+len(self.data)),options)

                self.workbook.close()
                self.result = int(not self.data==[])
            else:
                self.data=[]
                for dt,ms,eo,em,eu in self.unparsed_data:
                    if(dt>=fdate and dt<(tdate+datetime.timedelta(1)) and  eo==self.uid):
                        match=re.findall('.+=(.+?)\s',eu)
                        if match:
                            match=match[0]
                        else:
                            match=''
                        self.data.append((str(dt.date()),str(self.addSecs(dt.time(),ms)),em,match))
                if self.data:
                    options = {'banded_rows': 0, 'banded_columns': 1,'data': self.data,
                        'columns': [{'header':'Date'},
                                    {'header': 'Time'},
                                    {'header': 'Event Mess'},
                                    {'header': 'Event Userlooked'},
                                    ]} 
                    self.worksheet.add_table('B3:E'+str(3+len(self.data)),options)
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
  b=ParseData(datetime.datetime(2019,5,28),datetime.datetime(2019,5,30),1,2,False,'MOSG / 33KV / H03_CABLEFDR-H16 / MEASUREMENT / VOLTAGE VBR','All Signals','Average')
  #b=ParseData(datetime.datetime(2019,5,28),datetime.datetime(2019,7,6),1,1,True,'MOSG / 33KV / H38_39 BUS SEC / GIS SIGNALS / GIS VT  MCB TRIP','All Signals')
  print(b.result)
 