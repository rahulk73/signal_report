import pymysql.cursors
import datetime
import re
import datetime

class AccessDeniedError(Exception):
    pass
def generator(cursor,size=100):
    while True:
        rows=cursor.fetchmany(size)
        if not rows:
            break
        for row in rows:
            yield row
class GetSignalDataAnalog:
    def __init__(self,object_fullpath,signal_type,fdate,tdate,sh,un,pw,db):
        self.option={'All Signals':frozenset(['mappingsps','cs_voltageabsence_sps','cs_authostate_sps','cs_onoff_sps','cs_voltagerefpresence_sp','computedswitchpos_dps',
        'cs_busbarvchoice_sps','cs_acceptforcing_sps','groupsps','tapfct','cs_voltagepresence_sps','cs_closeorderstate_sps','userfunctionsps',
        'cs_voltagerefabsence_sps','moduledps','modulesps']),'All Controls': frozenset(['modulespc','switch_dpc','cs_ctrlonoff_spc','moduledpc']),
        'All Measurements':frozenset(['modulemv']),'All Metering':frozenset(['modulemeter'])}
        self.object_fullpath = object_fullpath
        self.result=()
        buffer_fdate = fdate + datetime.timedelta(minutes=-10)
        try:
            self.connection=pymysql.connect(host=sh,user=un,password=pw,db=db)
        except pymysql.err.OperationalError:
            raise AccessDeniedError
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
        except pymysql.err.OperationalError:
            raise AccessDeniedError
        except pymysql.err.ProgrammingError:
            self.result = None
        except Exception as e:
            print(e)
            self.result= -2
        finally:
            self.connection.close()

class GetSignalDataEvent:
    def __init__(self,fdate,tdate,object_fullpaths,sh,un,pw,db):
        self.result=()
        try:
            self.connection=pymysql.connect(host=sh,user=un,password=pw,db=db)
        except pymysql.err.OperationalError:
            raise AccessDeniedError
        try:
            self.result = None
            with self.connection.cursor() as cursor:
                if object_fullpaths:
                    self.sql = """
                    SELECT e.event_datetime,e.event_millisec,o.object_origin,object_description,e.event_mess,e.event_quality,e.event_userslogged
                    FROM events as e
                    inner join objects as o
                    on o.object_uid32=e.event_object_uid32
                    where (e.event_datetime between '{0}' and '{1}') and o.object_fullpath in """.format(fdate,tdate) + str(object_fullpaths)+' order by e.event_datetime'
                    print(self.sql)
                else:
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
    def __init__(self,sh,un,pw,db):
        self.result=()
        try:
            self.connection=pymysql.connect(host=sh,user=un,password=pw,db=db)
        except pymysql.err.OperationalError:
            raise AccessDeniedError
        try:
            with self.connection.cursor() as cursor:
                self.sql="SELECT object_fullpath,object_typ0,object_typ5,object_iecsignaladdr FROM objects order by object_fullpath"
                cursor.execute(self.sql)
                self.result=generator(cursor)
        except pymysql.err.OperationalError:
            raise AccessDeniedError
        except Exception as e:
            print(e)
            self.result=-2
        finally:
            self.connection.close()


if __name__ == "__main__":
    data = GetSignalDataAnalog(
        'MOSG / 33KV / H01_T40 TRF / MEASUREMENT / CURRENT IB',
        'All Measurements',
        datetime.datetime(2019,7,22,8,39),
        datetime.datetime(2019,7,22,8,55),
        sh='localhost',
        un='mcisadmin',
        pw='s$e!P!C!L@2014',
        db='pacis',
    )

    data = GetSignalDataAnalog(
        'MOSG / 33KV / H12_13 BUS SEC / Q0 CB / POSITION',
        'All Signals',
        datetime.datetime(2019,7,4,6,41),
        datetime.datetime(2019,7,4,7,0),
        sh='localhost',
        un='mcisadmin',
        pw='s$e!P!C!L@2014',
        db='pacis',
    )
    data = GetSignalDataAnalog(
        'MOSG / 33KV / H12_13 BUS SEC / Q0 CB / POSITION',
        'All Signals',
        datetime.datetime(2019,7,4,6,41),
        datetime.datetime(2019,7,4,7,0),
        sh='localhost',
        un='mcisadmin',
        pw='s$e!P!C!L@2014',
        db='pacis',
    )
    data = GetSignalDataEvent(
        datetime.datetime(2019,7,3,11,58),
        datetime.datetime(2019,8,3,12,00),
        ('MOSG / 33KV / H06_T10 (LV) INC / GIS SIGNALS / GIS CB SPRING DISCHARGED','MOSG / 33KV / H06_T10 (LV) INC / GIS SIGNALS / GIS DC SUPPLY FAIL'),
        sh='localhost',
        un='mcisadmin',
        pw='s$e!P!C!L@2014'
        ,db='pacis'
    )
    data = GetSignalDataEvent(
        datetime.datetime(2019,7,3,11,58),
        datetime.datetime(2019,8,3,12,00),
        (),
        sh='localhost',
        un='mcisadmin',
        pw='s$e!P!C!L@2014'
        ,db='pacis'
    )
    print(data.result)
    print(type(data.result))
    print(next(data.result))
        