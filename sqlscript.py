
import pymysql.cursors
import datetime
import re
from subprocess import run
def get_tables(uid):

    all_tables=[]
    try:
        file=open("pacis.sql",'r')
        pattern=re.compile(r"INSERT INTO `values_\d{1,}` VALUES [.\(\),0-9''-:\s]+?,"+str(uid))
        pattern2=re.compile(r'values_\d{1,}')
        for line in file:
            matches=pattern.finditer(line)
            for match in matches:
                sub_string=line[match.span()[0]:match.span()[1]]
                table=pattern2.findall(sub_string)
                all_tables.append(table[0])
    finally:
        file.close()
        return all_tables
def main(object_fullpath):
    result=()
    connection=pymysql.connect(host='localhost',user='mcisadmin',password='s$e!P!C!L@2014',db='pacis')
    try:
        with connection.cursor() as cursor:
            sql="select * from objects where object_fullpath='{}'".format(object_fullpath)
            cursor.execute(sql)
            result=cursor.fetchall()
            if result:
                run("mysqldump -u mcisadmin -ps$e!P!C!L@2014 pacis > pacis.sql",shell=True)
                all_tables=get_tables(result[0][1])
                run('del pacis.sql',shell=True)
                if all_tables:
                    sql="SELECT * FROM "+','.join(all_tables)
                    cursor.execute(sql)
                    return cursor.fetchall()
                else:
                    return 0
    except Exception as e:
        print(e)
        return -2
    finally:
        connection.close()
if __name__ == "__main__":
    main('MOSG / 11KV / K05_T40 LV INC / MEASUREMENT / VOLTAGE VYN')

