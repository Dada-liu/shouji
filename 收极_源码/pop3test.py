import sys,os
import re
import poplib
import  time

import email
from email.parser import Parser
from email.header import decode_header
#from email.utils import parseaddr


class mail():
    def __init__(self,pop3_server,username,password):
        self.pop3_server = pop3_server
        self.username = username
        self.password = password

    def decode_str(self,s):#字符编码转换
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value

    #获取邮件附件，
    def get_att(self,msg,path):
        """下载指定邮件的附件，返回所有附件的名称列表"""
        attachment_files = []
        for part in msg.walk():#.walk()方法为 os 模块中的函数，作用是递归输出层次结构，在这里递归输出MIME的层次结构
            file_name = part.get_filename()  # 获取附件名称类型
            #contType = part.get_content_type()#获取当前MIME层次的Content-Type
            if file_name:
                h = email.header.Header(file_name)
                #print('h',h)
                dh = email.header.decode_header(h)  # 将带有编码方式说明符的字符串分离成一个列表[(原始字符串，编码方式)]，
                #print('dh',dh)
                filename = dh[0][0]
                if dh[0][1]:
                    filename = self.decode_str(str(filename, dh[0][1]))  # 将附件名称可读化
                    print(filename)
                    # filename = filename.encode("utf-8")
                data = part.get_payload(decode=True)  # 下载附件
                att_file = open(path+ '\\' + filename, 'wb')  # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
                attachment_files.append(filename)#将附件名称加到列表中
                att_file.write(data)  # 保存附件
                att_file.close()
        return attachment_files



    def login(self):
        """获取pop3服务器名称、用户名、pop3密码，登录邮箱，返回邮箱设置的时间内(qq邮箱默认30天)的所有邮件数量"""
        try:
            server = poplib.POP3(self.pop3_server)
            server.set_debuglevel(1)
            server.getwelcome().decode('utf-8')
            server.user(self.username)
            server.pass_(self.password)
            mailnumber = server.stat()[0]#.stat()返回一个元组(x,y)，x为邮件数量，y为邮件占用空间；
            return [server,mailnumber]
        except:
            return 0

    def downloadfile(self,input_subject,path,index):
        server = poplib.POP3(self.pop3_server)
        server.set_debuglevel(1)
        server.user(self.username)
        server.pass_(self.password)
        #print("1",time.time())
        resp, lines, octets = server.retr(index)
        #print("2", time.time())
        try:
            content = b'\r\n'.join(lines).decode('utf-8')
        except UnicodeDecodeError:
            content = b'\r\n'.join(lines).decode('ISO-8859-1')
        #print("3", time.time())
        msg = Parser().parsestr(content)#把邮件内容转换为 message 对象
        #print("4", time.time())
        #对邮件主题进行判断
        mailsubject = msg["Subject"]
        mailsubject = email.header.decode_header(mailsubject)#将带有编码方式说明符的字符串分离成一个列表[(原始字符串，编码方式)]，
        #print("1", time.time())
        #有可能字符串并没有带编码方式，所以并不需要解码，直接可读
        if mailsubject[0][1] == None:
            mailsubject = mailsubject[0][0]
        else:
            mailsubject = mailsubject[0][0].decode(mailsubject[0][1])
        #print("邮件主题：",mailsubject)
        partten = re.compile(input_subject)
        mailsubject = partten.search(mailsubject)
        if mailsubject == None:
            return 0
        #print("正则表达式后的主题：",mailsubject.group())
        filelist = self.get_att(msg,path)
        #print('已下载附件：',filelist)
