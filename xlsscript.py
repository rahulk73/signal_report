import xlsxwriter
import datetime
import sqlscript
def roundTime(d1,roundTo):
    seconds=(d1-d1.min).seconds
    rounding=(seconds+roundTo/2)//roundTo*roundTo
    d2=d1+datetime.timedelta(0,rounding-seconds,-d1.microsecond)
    if(d2>d1):
        return d2
    else:
        return d2+datetime.timedelta(seconds=roundTo)
def addSecs(tm,ms):
    d1=datetime.datetime(1,1,1,tm.hour,tm.minute,tm.second)
    d1+=datetime.timedelta(milliseconds=ms)
    return d1.time()
def main(fdate,tdate,fq,radb,hidden,object_fullpath):
    try:    
        workbook = xlsxwriter.Workbook('./SignalLog/tables.xlsx')
        worksheet = workbook.add_worksheet()
        cell_format=workbook.add_format({'font_color':'white','bg_color':'#107896',\
    'font_name':'Times New Roman','bold':True,'underline':True,'center_across':True,'font_size':20})  
        caption = 'LOGS'
        worksheet.set_column('B:G', 20)
        worksheet.write('D1', caption,cell_format)
        unparsed_data=sqlscript.main(object_fullpath)
        if not unparsed_data:
            workbook.close()
            return -3
        data=[]
    
        if not hidden:
            option={1:1.0,2:60.0,3:3600.0}
            roundTo=(float)(fq)*option.get(radb)
            flag={}
            if unparsed_data:
                flag[unparsed_data[0][1]]=0
            for vm,vo,dt,ms,vq,vv in unparsed_data:
                dt=dt+datetime.timedelta(milliseconds=ms)
                if dt>tdate:
                    if flag[vo]==1:
                        flag[vo]=0
                    continue
                if dt<fdate :
                    continue
                elif dt >=fdate and dt < (tdate+datetime.timedelta(1)) and flag[vo]==0:
                    flag[vo]=1
                    cur_val=(vm,vo,vq,vv)
                    next_dt=roundTime(dt,roundTo)
                    data.append((str(dt.date()),str(addSecs(dt.time(),ms)),vm,vo,vq,vv))
                    continue
                elif flag[vo]==1 and dt<=next_dt:
                    cur_val=(vm,vo,vq,vv)
                elif flag[vo]==1 and dt>next_dt:
                    while dt>next_dt and next_dt<tdate:
                        data.append((str(next_dt.date()),str(next_dt.time()),*cur_val))
                        next_dt=roundTime(next_dt,roundTo)
                    cur_val=(vm,vo,vq,vv)
        else:
            for vm,vo,dt,ms,vq,vv in unparsed_data:
                if(dt>=fdate and dt<(tdate+datetime.timedelta(1))):
                    data.append((str(dt.date()),str(addSecs(dt.time(),ms)),vm,vo,vq,vv))
        if data:
            options = {'banded_rows': 0, 'banded_columns': 1,'data': data,
                'columns': [{'header':'Date'},
                            {'header': 'Time'},
                            {'header': 'Value Meantype'},
                            {'header': 'Value Object UID '},
                            {'header': 'Value Quality'},
                            {'header': 'Value Value'},
                            ]} 
            worksheet.add_table('B3:G'+str(3+len(data)),options)
        workbook.close()
        return int(not data==[])
    except PermissionError:
        return -1

if __name__ == "__main__":
  main(datetime.datetime(2019,5,28),datetime.datetime(2019,5,30),2,1,True,'MOSG / 33KV / H38_39 BUS SEC / BUSBAR PROT')
 