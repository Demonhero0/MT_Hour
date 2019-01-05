import xlrd
import xlwt
import os
from xlutils.copy import copy
import shutil
from total import total

def calTheLength(time):
    start,end = time.split("-")[0],time.split("-")[1]
    startList = start.split(":")
    endList = end.split(":")
    length = int(endList[0])-int(startList[0])+(int(endList[1])-int(startList[1]))/60
    return length

def add():

    date = input("date:")
    time = input("time:")
    localation = input("localation:")
    subject = "C语言"
    content = input("content:")
    evaluate = "好"
    MT = "苏健钟"
    students = input("students:")

    workbook = xlrd.open_workbook(u'model.xlsx')
    workbooknew = copy(workbook)
    ws = workbooknew.get_sheet(0)
    if date != "":
        ws.write(1, 1, date)
    if time != "":
        ws.write(1,3, time)
    if localation != "":
        ws.write(1,5, localation)
    if content != "":
        ws.write(3,1, content)
    if students != "":
        ws.write(9,4, students)
    ws.write(9,1, MT)
    ws.write(2,1, subject)
    ws.write(6,1, evaluate)
    workbooknew.save(u'model.xlsx')

    # workbook = xlrd.open_workbook(u'MT总工时.xlsx')
    # workbooknew = copy(workbook)
    # table = workbook.sheets()[0]
    # ws = workbooknew.get_sheet(0)
    # nrows = table.nrows
    # length = calTheLength(time)
    # ws.write(nrows,0,date)
    # ws.write(nrows,1,students)
    # ws.write(nrows,2,time)
    # ws.write(nrows,3,length)
    # workbooknew.save(u'MT总工时.xlsx')

    date = date.split("/")
    try:
        file_name = f'{date[0]}年{date[1]}月{date[2]}日 学业辅导打卡表'
    except:
        file_name = input("请输入打卡表日期:")
        file_name = f'{file_name} 学业辅导打卡表.xlsx'
    shutil.copy("model.xlsx", file_name)
    try:
        temp = file_name.split(" ")[0]
    except:
        temp = input("请输入文件夹的日期:")
    dir_name = f'{temp} 苏健钟小导师反馈'
    os.mkdir(dir_name)
    shutil.move(file_name,dir_name)

    total()
    print("Done")


if __name__ == "__main__":
    add()
