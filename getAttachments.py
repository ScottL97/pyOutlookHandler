# -*- coding:UTF-8 -*-
import os
import re
import sys
import time
import win32com.client
import win32timezone
from datetime import datetime

import xlsxHandler

downloadpath = os.path.abspath('.') + "\\attachments"
outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

def getmessages(accountname, foldertype):
    accounts = mapi.Folders
    messages = None
    for account in accounts:
        if account.Name == accountname:
            emailtypes = account.Folders # 收件箱、发件箱……
            for emailtype in emailtypes:
                if emailtype.Name == foldertype:
                    messages = emailtype.Items
    return messages

def getAttachments(accountname, subjectword, since, attype):
    cleanAttachments() # 清空attachments目录
    messages = getmessages(accountname, '收件箱') # 获取收件箱内容
    if messages != None:
        messages.Sort('[ReceivedTime]', True) # 按邮件接收日期排序
        tmpnum = 0
        for message in messages:
            # 筛选三天内的邮件
            if hasattr(message, 'ReceivedTime'):
                nowsec = time.time() # 现在的秒数（距1970-1-1）
                recvsec = time.mktime(datetime.strptime(str(message.ReceivedTime)[:10], '%Y-%m-%d').timetuple()) # 将收件日期转换成秒
                daysbetween = (nowsec - recvsec)/(24 * 60 * 60) # 距今天数
                # 筛选since天内的邮件
                if int(daysbetween) <= since: 
                    if hasattr(message, 'Subject'):
                        # 筛选主题带有subjectword的邮件
                        matchObj = re.search(r'.*' + subjectword + '.*', message.Subject) 
                        if matchObj != None:
                            # 获取附件列表
                            attachments = message.Attachments 
                            # 如果没有附件文件夹，创建文件夹
                            if not os.path.isdir(downloadpath): 
                                os.mkdir(downloadpath)
                            i = 1 # 每封邮件的附件序号从1开始
                            for attachment in attachments:
                                # 筛选attype格式的附件，因为邮件中的图片也会被当成附件
                                matchObj = re.match(r'.*\.' + attype + '$', attachment.FileName) 
                                if matchObj != None:
                                    att = attachments.Item(i)
                                    # 保存附件
                                    tmpfilename = downloadpath + '\\' + attachment.FileName[:-5] + '(' + \
                                        str(tmpnum) + ').' + attype
                                    print('[Sender]: %s' % message.SenderName)
                                    print('[Subject]: %s' % message.Subject)
                                    print('[Date]: %s' % str(message.ReceivedTime)[:10])
                                    print('[Since]: %s' % int(daysbetween) + "天前")
                                    print('[Attachment]: %s' % att)
                                    print('[SaveAs]: %s' % tmpfilename)
                                    print('-----------------------------------')
                                    att.SaveASFile(tmpfilename)
                                    tmpnum = tmpnum + 1
                                i = i + 1
                else:
                    break
    else:
        print('[Get 0 messages]')

def cleanAttachments():
    print('Cleaning attachments......\n')
    # 如果没有附件文件夹，创建文件夹
    if not os.path.isdir(downloadpath): 
        os.mkdir(downloadpath)
    for file in os.listdir(downloadpath):
        print('[Remove file]: %s' % file)
        os.remove(downloadpath + '\\' + file)
    print('-----------------------------------')


if __name__ == "__main__":
    if len(sys.argv) == 5:
        accountname = sys.argv[1] + '@inspur.com' # 邮箱账户名
        subjectword = sys.argv[2] # 邮件主题关键字
        since = int(sys.argv[3]) # 筛选since天的邮件
        attype = sys.argv[4] # 附件格式
    else:
        accountname = input('Please input your e-mail prefix: ') + '@inspur.com' # xxx@inspur.com
        subjectword = input('Please input your keyword: ') # such as '日志'
        since = int(input('Please input the days between the download start day and today: ')) # such as 10
        attype = input('Please input your download file type: ') # such as doc
    print('Getting attachments......\n')
    getAttachments(accountname, subjectword, since, attype)

