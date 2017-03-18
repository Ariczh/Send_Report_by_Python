#!/usr/bin/env python3
# -*- coding:utf-8 -*-
 
import os
import time
import datetime
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.utils import parseaddr, formataddr
import smtplib
import shutil
 
#遍历桌面，查找IT Daily Report开头的日期最晚的文件，请在目标文件夹不要放置太多杂物以免干扰
file = []
def getFiles(dir, keyword):
    res = []
    for root, directory, files in os.walk(dir):
        directory[:] = []#不搜索子文件夹，如需搜索请注释掉本句
        for filename in files:
            name, ext = os.path.splitext(filename)
            if keyword in name:
                res.append(os.path.join(filename))
    return res
file = getFiles('C:\\users\\zjy\\Desktop', 'IT Daily Report') #此处请该为自己的报告存放位置
fname = file[-1] #获取按文件名排序的最后一个文件
print (fname)
 
#时间函数
time = datetime.datetime.today()
date = str(time.strftime('20%y%m%d'))
 
#邮件地址编码器，因为存在BUG，实际没有用上
def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( Header(name, 'utf-8').encode(), addr.encode('utf-8') if isinstance(addr, unicode) else addr))
 
#邮件发送及接收信息，请酌情修改
your_mail = '*@#.com.cn'
your_pw = '***'
your_name = '张三'
to_name = '王五'
to_addr =  "*@#.com.cn"
cc_name = '李四'
cc_addr = '*@#.com.cn'
from_addr = your_mail
password = your_pw
smtp_server = 'c1.icoremail.net'
 
#邮件头部分
msg = MIMEMultipart()
msg['From'] = (u'%s <%s>' %(your_name, from_addr))
msg['To'] = (u'%s <%s>' % (to_name, to_addr))
msg['Cc'] = (u'%s <%s>' % (cc_name, cc_addr))
msg['Subject'] = Header(u'IT Daily Report周报 %s' % date,'utf-8').encode()
 
#邮件正文
msg.attach(MIMEText('见附件','plain','utf-8'))
 
#附件及增加的邮件头
with open ('/Users/zjy/Desktop/%s' % fname,'rb') as c:
    mime = MIMEBase('Sheet','xlsx', filename=fname)
    mime.add_header('Content-Disposition', 'attachment', filename=fname)
    mime.add_header('Content-ID', '<0>')
    mime.add_header('X-Attachment-Id', '0')
    # 把附件的内容读进来:
    mime.set_payload(c.read())
    encoders.encode_base64(mime)
    # 添加到MIMEMultipart:
    msg.attach(mime)
 
#合并发送和抄送地址
send_adds = []
send_adds.append(to_addr)
send_adds.append(cc_addr)
 
#发送部分
server = smtplib.SMTP(smtp_server,25)
#server.starttls() #禁用了SSL。如需启用请将上一句25改为465，会造成发送速度降低。
server.set_debuglevel(1)
server.login(from_addr, password)
server.sendmail(from_addr, send_adds, msg.as_string())
server.quit()
print ('文件%s已经作为附件发送至%s并抄送%s' % (fname,to_name,cc_name))
 
#已发送的文件移至历史文件夹
shutil.move('C:\\users\\zjy\\Desktop\\%s' % fname,'C:\\users\\zjy\\Desktop\\OUHSIS\\%s' % fname)
