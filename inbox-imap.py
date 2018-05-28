'''HR邮箱简历下载分类'''

import imaplib
import email
from email.parser import Parser
import xlwt
import xlrd
import os
import datetime


def get_email():    
    
    server = imaplib.IMAP4_SSL(host,993)

    server.login(username, password)
    # 去某指定文件夹
    res, data = server.select('JIRA')
    mailnum = data[0].decode()
    maxn = int(mailnum) - 1
    msg = []
    miss = 0  
 
    for i in range(maxn):
        j = str(int(mailnum) - i)
        try:
            res, msg_data = server.fetch(j, '(RFC822)')
            mssgg = email.message_from_bytes(msg_data[0][1])
            msg.append(mssgg)
        except:
            miss +=1
            continue

    server.logout()

    print('能读取邮件：' + str(len(msg)))
    
    return msg, miss  # 返回转化好的msg列表


def get_subject(message):
    subject = message.get('Subject')
    h = email.header.Header(subject)
    dh = email.header.decode_header(subject)
    chart = dh[0][1]
    if chart != None:
        # chart这个字符是该邮件的编码code
        head_line = dh[0][0].decode(chart)
    else:
        head_line = dh[0][0]
    return head_line


def get_sender(message):
    sender = message.get('From')
    sender_addr = email.utils.parseaddr(sender)
    return sender_addr[1]


def get_attach(message):
    parts = message.walk()
    for part in parts:
        fn = part.get_filename()       
        if fn:
            f = email.header.decode_header(fn)
            chat = f[0][1]
            if chat != None:
                fileName = f[0][0].decode(chat)
            else:
                fileName = f[0][0]
            name, sty = os.path.splitext(fileName)
            # 不下载附件中的图片格式，避免一些签名等无效图片
            if sty in ['.bmp', '.jpg', '.png', '.gif', '.tiff']:
                continue
            else:
                return(fileName)

def save_attach(message):
    parts = message.walk()
    for part in parts:
        fn = part.get_filename()
        
        if fn :
            f = email.header.decode_header(fn)
            chat = f[0][1]
            if chat!= None:
                # 同样处理编码问题
                fileName = f[0][0].decode(chat)
            else:
                fileName = f[0][0]
            # fileName = attr + fileName
            path = os.getcwd()
            abdir = os.path.join(path,  fileName)
            with open(abdir, 'wb') as file:
                data = part.get_payload(decode=True)
                file.write(data)
                file.close()

    return ()


def get_content(message):
    parts = message.walk()
    content = ''
    # 得到邮件所有编码相关类型
    chats = message.get_charsets()
    chat = ''
    for c in chats:
        if chat == '':
            if c != None:
                # 第一个code是邮件的编码code
                chat = c

    for part in parts:
        if content == '':
            if part.is_multipart() is False:

                content = part.get_payload(decode = True).decode(chat)
                
    return content


def get_date(message):
    date = message.get('Date')
    date_str = email.header.decode_header(date)
    dateTime = ''
    for t in date_str[0][0]: # 将邮件中提取的时间字段中有关时区的部分去掉
            if t != '+':
                dateTime += t
            elif t == '+':
                break
    date_time = datetime.datetime.strptime(dateTime, "%a, %d %b %Y %H:%M:%S ")
    # 将时间的str转化为可以直接比较时间的datetime格式
    return date_time



# 先判断是否为内推，然后判断校招（有无部门），剩下为社招（有无部门），最后是无法分类
# 问题是，校招中没有在标题中写明的，会被判断到社招中去。。。待解决


#读取要下载的文件的姓名list
namedata = xlrd.open_workbook('namelist.xls')
nametable = namedata.sheets()[0]
namelist = []
nrows = nametable.nrows
for r in range(nrows):
    namelist.append(nametable.row_values(r))

print(len(namelist))

#记录没有正常解析的邮件的相关情况
misslist = []
miss_sender = []

#初始化相关数据
host = 'imap.mxhichina.com'
username = 'yun.wang@aaa.com'
password = 'XXX'

msg_real = []
missn = 0

# 开始收取邮件
print("准备开始处理简历。。。")
msg_real, missn = get_email()


# 解析处理邮件
for mssg in msg:
    save_attach(mssg)








