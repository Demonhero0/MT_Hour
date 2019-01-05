import xlrd
import xlwt
import os
from xlutils.copy import copy
import shutil

def calTheLength(time):
    start,end = time.split("-")[0],time.split("-")[1]
    startList = start.split(":")
    endList = end.split(":")
    length = int(endList[0])-int(startList[0])+(int(endList[1])-int(startList[1]))/60
    return length


def total():

    dirs = os.listdir()
    total = list()
    for dir in dirs:
        if not os.path.isfile(dir):
            sub_dirs = os.listdir(dir)
            for file_name in sub_dirs:
                if '打卡表' in file_name:
                    goal_file = file_name
                    workbook = xlrd.open_workbook(f'{dir}/{goal_file}')
                    table = workbook.sheets()[0]
                    nrows = table.nrows
                    data = dict()
                    #    data['date'] = table.cell_value(1,1)
                    try:
                        date = xlrd.xldate_as_tuple(table.cell_value(1,1),0)
                        data['date'] = f'{date[0]}/{date[1]}/{date[2]}'
                    except:
                        data['date'] = table.cell_value(1,1)
                    data['time'] = table.cell_value(1,3)
                    data['students'] = table.cell_value(nrows-1,4)
                    data['length'] = calTheLength(data['time'])
                    total.append(data)
                    break
        else:
            continue

    workbook = xlrd.open_workbook(u'MT总工时.xlsx')
    workbooknew = copy(workbook)
    ws = workbooknew.get_sheet(0)
    ws.write(0,0,"日期")
    ws.write(0,1,"时间")
    ws.write(0,2,"学员")
    ws.write(0,3,"时长")
    ws.write(0,4,"总时长")
    total_time = 0.0
    i = 1
    print(total)
    for data in total:
        ws.write(i,0,data['date'])
        ws.write(i,1,data['time'])
        ws.write(i,2,data['students'])
        ws.write(i,3,data['length'])
        total_time += data['length']
        i += 1
    ws.write(1,4,total_time)
    print(f"total_time:{total_time}")
    workbooknew.save(u'MT总工时.xlsx')

if __name__ == "__main__":
    total()
