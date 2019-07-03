import pymysql.cursors
import datetime
import re
from subprocess import run
class GetSignalData:
    def __init__(self,object_fullpath):
        self.result=()
        self.connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
        try:
            with self.connection.cursor() as cursor:
                self.sql="select * from objects where object_fullpath='{}'".format(object_fullpath)
                cursor.execute(self.sql)
                self.result=cursor.fetchall()
                self.uid=self.result[0][1]
                if self.result:
                    run("mysqldump -u mcisadmin -ps$e!P!C!L@2014 pacis > pacis.sql",shell=True)
                    self.all_tables=self.get_tables(self.uid)
                    run('del pacis.sql',shell=True)
                    if self.all_tables:
                        self.sql="SELECT * FROM "+','.join(self.all_tables)
                        cursor.execute(self.sql)
                        self.result = cursor.fetchall()
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
                self.sql="SELECT object_fullpath,object_typ0,object_typ5 FROM objects order by object_fullpath"
                cursor.execute(self.sql)
                self.result=cursor.fetchall()
        except Exception as e:
            print(e)
            self.result=-2
        finally:
            self.connection.close()


if __name__ == "__main__":
    data=GetSignalData('MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN')
    print(data.result[0])

