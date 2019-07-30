import pymysql.cursors
import datetime
import re
import datetime
def generator(cursor,size=100):
    while True:
        rows=cursor.fetchmany(size)
        if not rows:
            break
        for row in rows:
            yield row
class GetSignalDataAnalog:
    def __init__(self,object_fullpath,signal_type,fdate,tdate):
        self.option={'All Signals':frozenset(['mappingsps','cs_voltageabsence_sps','cs_authostate_sps','cs_onoff_sps','cs_voltagerefpresence_sp','computedswitchpos_dps',
        'cs_busbarvchoice_sps','cs_acceptforcing_sps','groupsps','tapfct','cs_voltagepresence_sps','cs_closeorderstate_sps','userfunctionsps',
        'cs_voltagerefabsence_sps','moduledps','modulesps']),'All Controls': frozenset(['modulespc','switch_dpc','cs_ctrlonoff_spc','moduledpc']),
        'All Measurements':frozenset(['modulemv']),'All Metering':frozenset(['modulemeter'])}
        self.object_fullpath = object_fullpath
        self.result=()
        self.connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
        buffer_fdate = fdate + datetime.timedelta(minutes=-10)
        try:
            self.result = None
            with self.connection.cursor() as cursor:
                self.sql = "select object_typ5,object_uid32 from objects where object_fullpath='{}'".format(object_fullpath)
                cursor.execute(self.sql)
                self.rows = cursor.fetchall()
                self.uid=None
                for row in self.rows:
                    if row[0] in self.option[signal_type]:
                        self.uid=row[1]
                        break   
                if self.uid:
                    if signal_type in ['All Measurements','All Metering']:
                        num = int(str(self.uid)[-2:])
                        self.sql="SELECT * FROM values_"+str(num)+" where value_object_uid32 = {0} and value_datetime between '{1}' and '{2}'".format(self.uid,buffer_fdate,tdate)
                        cursor.execute(self.sql)
                        self.result=generator(cursor)
                    else:
                        self.sql="""select event_datetime,event_millisec,event_mess,event_userslogged
                        from events
                        where event_object_uid32={0} and event_datetime between '{1}' and '{2}'
                        """.format(self.uid,fdate,tdate)
                        cursor.execute(self.sql)
                        self.result=generator(cursor)
        except pymysql.err.ProgrammingError:
            self.result = None
        except Exception as e:
            print(e)
            self.result= -2
        finally:
            self.connection.close()

class GetSignalDataEvent:
    def __init__(self,fdate,tdate):
        self.result=()
        self.connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
        try:
            self.result = None
            with self.connection.cursor() as cursor:
                self.sql = """
                SELECT e.event_datetime,e.event_millisec,o.object_origin,object_description,e.event_mess,e.event_quality,e.event_userslogged
                FROM events as e
                inner join objects as o
                on o.object_uid32=e.event_object_uid32
                where e.event_datetime between '{0}' and '{1}'
                """.format(fdate,tdate)
                cursor.execute(self.sql)
                self.result = generator(cursor)
        except pymysql.err.ProgrammingError:
            self.result = None
        except Exception as e:
            print(e)
            self.result= -2
        finally:
            self.connection.close()
class GetSignals:
    def __init__(self):
        self.result=()
        self.connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
        try:
            with self.connection.cursor() as cursor:
                self.sql="SELECT object_fullpath,object_typ0,object_typ5,object_iecsignaladdr FROM objects order by object_fullpath"
                cursor.execute(self.sql)
                self.result=generator(cursor)
        except Exception as e:
            print(e)
            self.result=-2
        finally:
            self.connection.close()


if __name__ == "__main__":
    #data=GetSignalData('MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN','All Signals')
    # data=GetSignalData('MOSG / 33KV / H32_T30 (LV) INC / BCU SYNCROCHECK / ON/OFF SPS','All Signals')
    #data=GetSignals()
    # data=GetSignalDataAnalog('MOSG / 33KV / H01_T40 TRF / MEASUREMENT / CURRENT IB','All Measurements',datetime.datetime(2019,7,22,8,39),datetime.datetime(2019,7,22,8,55))
    # data=GetSignalDataAnalog('MOSG / 33KV / H12_13 BUS SEC / Q0 CB / POSITION','All Signals',datetime.datetime(2019,7,4,6,41),datetime.datetime(2019,7,4,7,0))
    data = GetSignalDataEvent(datetime.datetime(2019,6,19,11,58),datetime.datetime(2019,6,19,12,00))
    print(data.result)
    print(type(data.result))
    print(next(data.result))
    # f = open('output.txt','w')
    # try:
    #     while True:
    #         f.write(str(next(data.result))+'\n')
    # except StopIteration:
    #     print('stopped')
    # except TypeError:
    #     print('Result is empty')
    # finally:
    #     f.close()
        