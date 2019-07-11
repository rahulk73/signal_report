import pymysql.cursors
import datetime
import re
from subprocess import run
def generator(cursor,size=100):
    while True:
        rows=cursor.fetchmany(size)
        if not rows:
            break
        for row in rows:
            yield row
class GetSignalData:
    def __init__(self,object_fullpath,signal_type):
        self.option={'All Signals':frozenset(['mappingsps','cs_voltageabsence_sps','cs_authostate_sps','cs_onoff_sps','cs_voltagerefpresence_sp','computedswitchpos_dps',
        'cs_busbarvchoice_sps','cs_acceptforcing_sps','groupsps','tapfct','cs_voltagepresence_sps','cs_closeorderstate_sps','userfunctionsps',
        'cs_voltagerefabsence_sps','moduledps','modulesps']),'All Controls': frozenset(['modulespc','switch_dpc','cs_ctrlonoff_spc','moduledpc']),
        'All Measurements':frozenset(['modulemv']),'All Metering':frozenset(['modulemeter'])}
        self.result=()
        self.connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
        try:
            self.result = 0
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
                    if signal_type  == 'All Measurements':
                        self.sql="SELECT * FROM "+str(self.uid)[-2:]
                        cursor.execute(self.sql)
                        self.result=generator(cursor)
                    else:
                        self.sql="select event_datetime,event_millisec,event_object_uid32,event_mess,event_userslogged from `events` where event_object_uid32={}".format(self.uid)
                        cursor.execute(self.sql)
                        self.result=generator(cursor)
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
    data=GetSignalData('MOSG / 33KV / H32_T30 (LV) INC / BCU SYNCROCHECK / ON/OFF SPS','All Signals')
    #data=GetSignals()
    print(data.result)
    print(next(data.result))