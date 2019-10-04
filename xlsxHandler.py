# -*- coding:UTF-8 -*-
import xlrd
import openpyxl # 经测试，xlwt无法保存表格为xlsx格式
import os
import re
import time
from xlutils.copy import copy

attachmentsdir = os.path.abspath('.') + '\\attachments'

# 对表格单元格的几种操作
ADD = 0
MODIFY = 1
DELETE = 2

def write_cell(modelpath, rownum, colnum, value):
    modelbooknew = openpyxl.load_workbook(modelpath) # 获取模板表格文件
    modelsheetnew = modelbooknew.active # 获取工作表
    # openpyxl中行和列从1开始
    modelsheetnew.cell(row = rownum, column = colnum).value = value
    modelbooknew.save(modelpath) # 覆盖掉原文件

def process_xlsx(modelpath, xlsx):
    print(xlsx)
    modelbook = xlrd.open_workbook(modelpath) # 待修改/填充的模板表格文件
    workbook = xlrd.open_workbook(attachmentsdir + '\\' + xlsx) # 要比较的表格文件
    # 默认处理第一个sheet
    modelsheet = modelbook.sheet_by_index(0)
    worksheet = workbook.sheet_by_index(0) 
    print('行数：%d，列数：%d' % (worksheet.nrows, worksheet.ncols))
    if worksheet.nrows != modelsheet.nrows or worksheet.ncols != modelsheet.ncols:
        return -1
    # 表头
    # header = worksheet.row_values(0) 
    # print(header)
    for i in range(1, worksheet.nrows):
        # 只处理模板中工号不为空的行
        if modelsheet.row_values(i)[0] != "":
            # 当单元格内容不同时，添加/修改/删除模板中的相应值
            if modelsheet.row_values(i)[14] != worksheet.row_values(i)[14]:
                if modelsheet.row_values(i)[14] == "":
                    write_cell(modelpath, i + 1, 15, worksheet.row_values(i)[14])
                    addlog(modelsheet.row_values(i)[2], worksheet.row_values(i)[14], None, \
                        modelsheet.row_values(i)[3], ADD)
                elif worksheet.row_values(i)[14] != "":
                    write_cell(modelpath, i + 1, 15, worksheet.row_values(i)[14])
                    addlog(modelsheet.row_values(i)[2], modelsheet.row_values(i)[14] + '|' + \
                        worksheet.row_values(i)[14], None, modelsheet.row_values(i)[3], MODIFY)
            if modelsheet.row_values(i)[15] != worksheet.row_values(i)[15]:
                if modelsheet.row_values(i)[15] == "":
                    write_cell(modelpath, i + 1, 16, worksheet.row_values(i)[15])
                    addlog(modelsheet.row_values(i)[2], None, worksheet.row_values(i)[15], \
                        modelsheet.row_values(i)[3], ADD)
                elif worksheet.row_values(i)[15] != "":
                    write_cell(modelpath, i + 1, 16, worksheet.row_values(i)[15])
                    addlog(modelsheet.row_values(i)[2], None, modelsheet.row_values(i)[15] + '|' + \
                        worksheet.row_values(i)[15], modelsheet.row_values(i)[3], MODIFY)

def doxlsxHandler():
    modelpath = os.path.abspath('.')
    # 获取表格模板文件
    for file in os.listdir(modelpath):
        matchObj = re.match(r'.*\.xlsx', file)
        if matchObj != None:
            modelpath = modelpath + '\\' + file
            break
    # 遍历附件目录
    for file in os.listdir(attachmentsdir):
        matchObj = re.match(r'.*\.xlsx', file)
        if matchObj != None:
            if process_xlsx(modelpath, file) == -1:
                print('[Wrong file]: ' + file)
            print('--------------------------------------------------------------------------------------')

def addlog(name, explaination, comment, date, logtype):
    # 创建日志目录
    if not os.path.isdir('logs'):
        os.mkdir('logs')
    # 操作记录变量
    record = ""
    # 格式化操作记录
    if logtype == ADD:
        if explaination != None:
            record = '[%s ADD Explaination: %s]: %s' % (name, date, explaination)
        if comment != None:
            record = '[%s ADD Comment: %s]: %s' % (name, date, comment)
    elif logtype == MODIFY:
        if explaination != None:
            print('[%s MODIFY Explaination: %s]: %s - %s' % (name, date, explaination.split('|')[0], \
            explaination.split('|')[1]))
        if comment != None:
            print('[%s MODIFY Comment: %s]: %s - %s' % (name, date, comment.split('|')[0], \
            comment.split('|')[1]))
    elif logtype == DELETE:
        if explaination != None:
            print('[%s DELETE Explaination: %s]: %s' % (name, date, explaination))
        if comment != None:
            print('[%s DELETE Comment: %s]: %s' % (name, date, comment))

    if record != "":
        # 输出操作记录
        print(record)
        # 写入日志文件
        with open('logs\\' + time.strftime('%Y/%m/%d %H:%M:%S', time.localtime()) + '.log', 'a', encoding='utf-8') as log:
            log.write(record + '\n')

if __name__ == '__main__':
    print('\nHandling attachments......\n')
    doxlsxHandler()