
import sys
import os
import getAttachments
import xlsxHandler

if __name__ == "__main__":
    if len(sys.argv) == 5:
        accountname = sys.argv[1] + '@inspur.com' # 邮箱账户名
        subjectword = sys.argv[2] # 邮件主题关键字
        since = int(sys.argv[3]) # 筛选since天的邮件
        attype = sys.argv[4] # 附件格式
    elif os.path.isfile(os.path.abspath('.') + "\\params.txt"):
        with open(os.path.abspath('.') + "\\params.txt", 'r', encoding='utf-8') as f:
            for line in f:
                params = line.split(' ')
                accountname = params[0] + '@inspur.com'
                subjectword = params[1]
                since = int(params[2])
                attype = params[3]
    else:
        accountname = input('Please input your e-mail prefix: ') + '@inspur.com' # xxx@inspur.com
        subjectword = input('Please input your keyword: ') # such as '日志'
        since = int(input('Please input the days between the download start day and today: ')) # such as 10
        attype = input('Please input your download file type: ') # such as doc
    print('Getting attachments......\n')
    getAttachments.getAttachments(accountname, subjectword, since, attype)
    print('\nHandling attachments......\n')
    xlsxHandler.doxlsxHandler()
    print('\nFinished! Hava a nice day!\n')
    input('Press enter to exit...')
