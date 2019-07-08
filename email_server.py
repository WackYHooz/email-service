#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# !/usr/bin/python
# -*- coding:UTF-8 -*-
from email.parser import Parser
from email.header import decode_header
import imaplib, string, email
import os
import poplib
from urllib import parse
import requests
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import HTMLParser
import time
import sys
import re

def validateTitle(title):
    rstr = r"[\/\\\:\*\?\"\<\>\|]"  # 非法字符，包含: '/ \ : * ? " < > |'
    new_title = re.sub(rstr, "_", title)  # 替换为下划线
    return new_title

def mkdir(path):
#     去除首位空格
    path=path.strip()
#     去除尾部 \ 符号
    path=path.rstrip("\\")
#     判断路径是否存在
    isExists=os.path.exists(path)
#     如果不存在则创建目录
    if not isExists:
        os.makedirs(path)  
        print(path+' Success to build the path!')
        return True
#     如果目录存在则不创建，并提示目录已存在
    else:
        print(path+' The Dir is exsist!')
        return False

# 邮件的Subject或者Email中包含的名字都是经过编码后的str，要正常显示，就必须decode
def decode_str(s):
    value, charset = decode_header(s)[0]
    # decode_header()返回一个list，因为像Cc、Bcc这样的字段可能包含多个邮件地址，所以解析出来的会有多个元素。上面的代码我们偷了个懒，只取了第一个元素。
    if charset:
        value = value.decode(charset)
#         value = value.decode('gbk')
    return value

# 判断是否有附件，有则存储在指定路径下（默认为当前文件夹下）
def find_atchmt(msg, sender, Dirpath=sys.path[0]+'/'+'tmp_attchment'):
    flag = 0# 判断是否有附件
    for par in msg.walk():
        name = par.get_filename()
        if not name is None:
            name = decode_str(name)
#             使附件名称规范化，防止无法保存
            name = validateTitle(name)
            sender = validateTitle(sender)
            try:
                mkdir(Dirpath)
            except Exception as e:
                print('Fail to make dir: %s' %e)
                print('-'*50)
                return False
            fpath = Dirpath+'/'+sender+'_'+name
            data = par.get_payload(decode=True)
            try:
                f = open(fpath, 'wb')# 注意一定要用wb来打开文件，因为附件一般都是二进制文件
                f.write(data)
                f.close()
                print('save to file %s succeed'%fpath)
                print('-'*50)
                flag = 1
            except Exception as e:
                print('open file name error: %s' %e)
                print('-'*50)
                return False
    if flag == 0:        
        print('No attachment in this email.')
        print('-'*50)
        return None
    else:
        return True

# 解析邮件的body
def print_info(msg,indent=0):
#     print(type(msg), msg)
    if indent ==0:
        for header in ["From", "Date", "To", "Subject", "name"]:
#             print(header)
            value = msg.get(header,"")
#             print(value, type(value))
#             if header == "Subject" and value is None:
#                 continue
            if value or (header == "Subject" and value is ""):
                if header == "From":
                    sender = decode_header(value)
                    try:
                        tmp = []
                        for i in sender:
                            tmp.append(i[0])
#                         print('tmp:', tmp)
                        lst_tmp = []
                        for i in tmp:
    #                         print(i)
                            try:
                                lst_tmp.append(i.decode('gbk'))
                            except Exception as e:
#                                 print('Fail to decode the name of the sender')
                                lst_tmp.append(i.decode())
                        sender = "".join(lst_tmp)
                        sender = sender.replace("\"", "")
                        bgn = sender.find("<", 0)
                        if bgn != -1:
                            sender = sender[:bgn-1]+sender[bgn:]
#                             print(sender)
                    except Exception as e:
                        print('No need to transform')
                        sender = sender[0][0]
                if header == "Date":
                    date = decode_header(value)[0][0]
#                     print('Date: ', Date[0])
                if header == "Subject":
                    if not value is "":
                        value = decode_str(value)
#                         print('Subject: ', value)
#                     if value == "邮件主题名" :#取18封邮件中需要的邮件
                    else:
                        print('The Subject is None!')
                    subject = value
                    for par in msg.walk():
#                         name = par.get_filename()
#                         if not name is None:
#                             print('get file! Name is %s' %name)
#                         else:
#                             print('no file in this email!')
#                         print('name_1: ', name_1)
                        if not par.is_multipart(): # 这里要判断是否是multipart，是的话，里面的数据是无用的，至于为什么可以了解mime相关知识。
                            content_type = par.get_content_type()
#                             print(content_type)

                            if content_type == 'text/plain' or content_type == 'text/html':
#                                 print(par,type(par))
                                content = par.get_payload(decode=True)
                                try:
                                    content = content.decode('utf-8')
                                except Exception as e:
                                    print('Fail to use \'utf-8\' to decode! use \'gbk\' to try again...')
                                    content = content.decode('gbk')
                                if content_type == 'text/html':
                                    print('Error! This is a html email, which isn\'t accepted for the moment.')
#                                     content = HTMLParser(content)
                                    return None, None, None, None
                                return sender, date, subject, content

# 获取邮件的有效信息
def getMail(host, imap, user, passwd, pop_port=993, imap_port=110):
    contents = []
    senders = []
    subjects = []
    
    try:
        print('Trying to connect to IMAP_SLL: %s | Port ID: %s' %(imap, imap_port))
        M = imaplib.IMAP4_SSL(imap, imap_port)
    except Exception as e:
        print('Failed to access: %s' %e)
        return False
#         sys.exit()
    try:
        print('Trying to connect to POP3: %s | Port ID: %s' %(host, pop_port))
        p = poplib.POP3(host, pop_port)
    except Exception as e:
        print('Failed to access: %s' %e)
        return False
#         sys.exit()
    try:
        print('Login...')
        p.user(user)
        p.pass_(passwd)
        M.login(user,passwd)
#     except (poplib.error_proto,e):
    except Exception as e:
        print ("Login failed: %s" %e)
        return False
#         sys.exit()
    try:
        print('try to connect to STMP...')
        smtpObj = smtplib.SMTP()
        smtpObj.connect(SMTP_host, SMTP_port)
        smtpObj.login(user, passwd)
    except Exception as e:
        print('Failed to access: %s' %e)
        return False
#     print('Success to connect!')
    
    try:
#         获取未读邮件的UID
        try:
            M.select('INBOX', False)
            _, data = M.search(None, 'unseen')
            lst = data[0].decode(encoding='utf-8').split(" ")# 对btye编码进行解码，获取str形式的UID
#             print(lst)
            lst = list(map(str, lst))# 将UID转换为list
            if data[0] == b'':
                print('No new unseen mail here. Continue to scan after %d(s)' %wait_time)
                print('_'*50)
                return None, None, None
            print('%d unseen email(s) here.' %int(len(lst)))
            print('-'*50)
        except Exception as e:
            print('Unknown error: %d' %e)
            return False
#             sys.exit()
#         lst = [0,1]
        for index in lst:
#             index = str(index)
            typ, data = M.fetch(index, '(RFC822)')
#             print('typ: ', typ, data)
            try:
                text = data[0][1].decode('utf-8')
            except Exception as e:
                print('Failed to decode by \'utf-8\', try to decode by \'gbk\'...')
                try:
                    text = data[0][1].decode('gbk')
                except Exception as e:
                    print('Failed to decode: %s' %e)
                    sys.exit()
#             print('text: ', text, type(text))
            msg = email.message_from_string(text)
#             print('Try to get info of this email...')
            sender, date,  subject, ori_content = print_info(msg)# 获取发件人，主题和正文文本
            if sender is None:
                continue
            content = del_Fw(ori_content)
#             print(content)
            print('Telling with signature...')
            content = del_signature(content)
            print('Finding attachment(s)...')
            find_atchmt(msg, sender)
#             print('Success!')
            try:
#             if not senders is None:
                query = content
                body = {"userId": userId, "botId": botId, "query": query, "type": _type}
#                 print('body: ', body)
                r = requests.post(URL, json=body)
            #     print(r.text)
                if r.status_code == 200:
                    reply_bgn = r.text.find('\"reply\": ',0)
                    reply_end = len(r.text)-2
                    reply = r.text[reply_bgn+10:reply_end].encode('utf-8').decode('unicode_escape')
                else:
                    print('Failed to request API: ', r)
                    return False
            except Exception as e:
                print('Fail to request the API: %s' %e)
#             尝试发送邮件
            try:
                from_addr, to_addr, message = send_email(username, sender, subject, reply)
                smtpObj.sendmail(from_addr, to_addr, message)
                print("Success to send the email!")
            except smtplib.SMTPException:
                print('Error! Fail to send the email: %s' %e)
                return False
#             尝试转发邮件
            try:
                to = sender_name+'<%s>' % username
                Fw_content = '''
%s\r\n\r\n
------------------ Original ------------------\r\n
From: %s;\r\n
Date: %s\r\n
To: %s;\r\n
Subject: %s\r\n\r\n
%s''' %(reply, sender, date, to, subject, ori_content)
#                 如果需要转发给多人，则循环多次进行发送
                if isinstance(Fw_email, list):
                    for i in range(len(Fw_email)):
                        print(Fw_email)
                        Fw_addr = Fw_name[i]+'<%s>' %Fw_email[i]
                        from_addr, to_addr, message = send_email(username, Fw_addr, subject, Fw_content)
                        smtpObj.sendmail(from_addr, to_addr, message)
                        print("Success to forward the email!")
                else:
                    Fw_addr = Fw_name[i]+'<%s>' %Fw_email[i]
                    from_addr, to_addr, message = send_email(username, Fw_addr, subject, Fw_content)
                    smtpObj.sendmail(from_addr, to_addr, message)
                    print("Success to forward the email!")
            except smtplib.SMTPException:
                print('Error! Fail to forward the email: %s' %e)
                return False
            try:
                M.store(str(index), '+FLAGS', '(\\Seen)')
                print('This \"unseen\" email haw been set to \"seen\"')
                print('_'*50)
            except Exception as e:
                print(e)
                print('Unknown Error! Can\'t set this mail to the status of \'Seen\'\nTo prevent users from receiving repeated replies, we skip this mail.')
#                 senders.pop()
#                 contents.pop()
        return True
            
    except Exception as e:
        print ('Error: %s' % e)
        return False

# 去签名
def del_signature(content):
#     读取常见的正文与签名分界处格式
    try:
        sig_format = [line.strip() for line in open("email_server_signature.ini",encoding='gb18030',errors='ignore').readlines()]
    except Exception as e:
        print('Can\'t find \"email_server_signature.ini\" here so we can\'t tell with the signature.')
        return content
#     查找最靠近文末的疑似签名处
    max_bgn = -1
    bgn = 0
    for elmt in sig_format:
        while bgn != -1:
            elmt = elmt.encode('utf-8').decode('unicode_escape')
            elmt_len = len(elmt.split("\r\n"))
            bgn = content.find(elmt, bgn)# 对转义格式进行重解码
            if not bgn == -1:
                if bgn > max_bgn:
                    max_bgn = bgn
                bgn += 1
        bgn = 0
    if max_bgn == -1:
        print('Seems like no signature in this email.')
    else:
        sig = content[max_bgn:]
        sig_lst = sig.split("\r\n")
        max_len = 0
        flag = 0# 判断是否为签名的标识
#         获取疑似签名文本中最长的字符数
        for line in sig_lst:
            if len(line)>max_len:
                max_len = len(line)
#         最长字符串小于指定阈值，则标识+1
        if max_len < 35 or len(sig_lst) >= elmt_len+3:
            print('The content below seems like a signature:\n%s' % content[max_bgn:])
            content = content[:max_bgn]
        else:
            print('Seems like no signature in this email.')
    print('-'*50)    
    return content

# 删除正文中的转发邮件的Title
def del_Fw(content):
    bgn = 0
    max_bgn = -1
#     查找最后一处转发邮件的Title
    while bgn != -1:
        bgn = content.find('------------------ Original ------------------', bgn)
#         保存最后一处Title
        if bgn > max_bgn:
            max_bgn = bgn
#         找到字符串后，若不+1则会死循环
        if bgn != -1:
            bgn += 1
    if max_bgn == -1:
        return content
    else:
        print('This is a Fw email.')
        bgn = content.find('Subject:', max_bgn)
        bgn = content.find('\n', bgn)
        print('This content seems like the header of the Fw email:\n%s' %content[:bgn+8])
        content = content[bgn+8:]
        print('-'*50)
        return content

# 根据给定邮箱和和指定内容，发送邮件
def send_email(from_addr, to_addr, subject, msg):
    message = MIMEText(msg, 'plain', 'utf-8')# 文本
    message['From'] = Header((sender_name+'<%s>') %from_addr, 'utf-8')# 发送方邮箱
    message['Sender'] = Header(sender_name, 'utf-8')# 发送者
    message['To'] =  Header(to_addr, 'utf-8')# 接收方邮箱
#         message['From'] = sender_name# 发送方邮箱
#         message['To'] =  to_addrs[i]# 接收方邮箱
    Date = time.strftime("%a,%d %b %Y %H:%M:%S %z")
    message['Date'] = Date
    subject = 'reply: '+subject
    message['Subject'] = Header(subject, 'utf-8')
    print('-'*11+'Preview the email to be sent'+'-'*11)
    print('Subject: %s\nDate: %s\nTo: %s\nContent: \n%s' % (subject, Date, to_addr, msg))
    print('-'*50)
    return from_addr, to_addr, message.as_string()


if __name__ == '__main__':
#     加载预设变量
    print('Loading Preset variables...')
    ini_param_tmp = [line.strip() for line in open("email_server.ini",encoding='gb18030',errors='ignore').readlines()]
    ini_param = []
    for line in ini_param_tmp:
        line = line.replace("\t", "").                    replace("=", "").replace("：", "")
        ini_param.append(line)
#     print(ini_param)
    pop_host = str(ini_param[1][6:])# pop服务器
    pop_port = int(ini_param[2][8:])# pop端口号
    imap_host = str(ini_param[3][7:])# imap服务器
    imap_port = int(ini_param[4][9:])# imap端口号
    SMTP_host = str(ini_param[5][7:])# 发送邮件服务器
    SMTP_port = int(ini_param[6][7:])# 发送邮件端口号
    username = str(ini_param[7][6:])# 邮箱用户名
    sender_name = str(ini_param[8][5:])# 发送者名称
    password = str(ini_param[9][5:])# 邮箱密码
    Fw_email = str(ini_param[10][5:]).split(',')
    Fw_name = str(ini_param[11][5:]).split(',')
    URL = str(ini_param[14][7:])# Email_service接口
    userId = str(ini_param[15][6:])# 用户ID 参考Email_service接口文档
    botId = str(ini_param[16][5:])# 参考Email_service接口文档
    _type = str(ini_param[17][5:])# 参考Email_service接口文档
    wait_time = int(ini_param[20][7:])# 检测间隔时间
    print('Success loading!')

#     定期循环检查新邮件
    while 1:
        status = getMail(pop_host, imap_host, username, password, pop_port, imap_port)
#         print(senders, contents)
        time.sleep(wait_time)

