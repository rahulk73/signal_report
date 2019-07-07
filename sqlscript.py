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
        self.option={'All Signals':1,'All Controls':0}[signal_type]
        self.result=()
        self.connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
        try:
            with self.connection.cursor() as cursor:
                self.sql="select * from objects where object_fullpath='{}'".format(object_fullpath)
                cursor.execute(self.sql)
                self.result=cursor.fetchall()
                row=()
                for row in self.result:
                    if row[42]:
                        schar=row[42][-1]
                        if schar != 'c' and self.option == 1:
                            self.uid=row[1]
                            break
                        elif schar == 'c' and self.option == 0:
                            self.uid=row[1]
                            break
                    elif self.option == 1:
                        self.uid=row[1]
                        break
                if self.result and row:
                    run("mysqldump -u mcisadmin -ps$e!P!C!L@2014 pacis > pacis.sql",shell=True)
                    self.all_tables=self.get_tables(self.uid)
                    run('del pacis.sql',shell=True)
                    if self.all_tables:
                        self.sql="SELECT * FROM "+','.join(self.all_tables)
                        cursor.execute(self.sql)
                        self.result=generator(cursor)
                    else: 
                        self.result = 0
        except Exception as e:
            print(e)
            self.result= -2
        finally:
            self.connection.close()
    def get_tables(self,uid):
        self.all_tables=[]
        try:
            self.file=open("pacis.sql",'r')
            pattern=re.compile(r"INSERT INTO `values_\d{1,}` VALUES [.\(\),0-9''-:\s]+?,"+str(uid))
            pattern2=re.compile(r'values_\d{1,}')
            for line in self.file:
                matches=pattern.finditer(line)
                for match in matches:
                    sub_string=line[match.span()[0]:match.span()[1]]
                    table=pattern2.findall(sub_string)
                    self.all_tables.append(table[0])
        finally:
            self.file.close()
            return self.all_tables


class GetSignals:
    def __init__(self):
        self.result=()
        self.connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
        try:
            with self.connection.cursor() as cursor:
                self.sql="SELECT object_fullpath,object_typ0,object_typ5 FROM objects where (object_typ5='' and object_typ0='scs')or(object_typ5<>'') order by object_fullpath"
                cursor.execute(self.sql)
                self.result=generator(cursor)
        except Exception as e:
            print(e)
            self.result=-2
        finally:
            self.connection.close()


if __name__ == "__main__":
   # data=GetSignalData('MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN','All Signals')
   # data=GetSignalData('MOSG / 33KV / H06_T10 (LV) INC / CONTROL MODE / CONTROL MODE','All Controls')
    data=GetSignals()
    print(data.result)

