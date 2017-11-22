# -*- coding:utf-8 -*-

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart,MIMEBase
from email import encoders
from email.header import Header
from email.utils import parseaddr,formataddr
import smtplib

def _format_addr(s):
    name,addr=parseaddr(s)
    return formataddr((Header(name,'utf-8').encode(),addr))


def formContent(msg):
    #添加文本
    #发送纯文本内容
    #msg=MIMEText('Hello,send by Python...','plain','utf-8')
    #发送html格式的正文
    msg_text = MIMEText('<html><body><h1>Hello</h1>' +
        '<p>send by <a href="http://www.python.org">Python</a>...</p>' +
        '<p><img src="cid:0"></p>' +
        '</body></html>', 'html', 'utf-8')
    msg.attach(msg_text)
    #添加附件
    with open('C:\Users\...\Pictures\Devops.png','rb') as f:
        mime=MIMEBase('image','png',filename='test.png')
        # 加上必要的头信息:
        mime.add_header('Content-Disposition', 'attachment', filename='test.png')
        mime.add_header('Content-ID', '<0>')
        mime.add_header('X-Attachment-Id', '0')
        # 把附件的内容读进来:
        mime.set_payload(f.read())
        # 用Base64编码:
        encoders.encode_base64(mime)
        # 添加到MIMEMultipart:
        msg.attach(mime)

    with open('C:\Users\...\Pictures\Template.xlsx','rb') as f:
        mime=MIMEBase('file','xlsx',filename='test.xlsx')
        # 加上必要的头信息:
        mime.add_header('Content-Disposition', 'attachment', filename='test.xlsx')
        mime.add_header('Content-ID', '<1>')
        mime.add_header('X-Attachment-Id', '1')
        # 把附件的内容读进来:
        mime.set_payload(f.read())
        # 用Base64编码:
        encoders.encode_base64(mime)
        # 添加到MIMEMultipart:
        msg.attach(mime)

if __name__=='__main__':
    from_addr=input('From:')
    password=input('Password:')
    to_addr=input('To:')
    smtp_server=input('SMTP server:')
    msg=MIMEMultipart()
    #发送带附件的邮件
    #创建一个组合附件，包括文本和附件
    msg['From']=_format_addr('Python Lover <%s>' % from_addr)
    msg['To']=_format_addr('Administrator <%s>' % to_addr)
    msg['Subject']=Header('Python发送测试邮件','utf-8').encode()
    formContent(msg)
    
    server=smtplib.SMTP(smtp_server,25)
    #使用加密
    #server.starttls()
    server.set_debuglevel(0)
    server.login(from_addr,password)
    server.sendmail(from_addr,[to_addr],msg.as_string())
    server.quit()
