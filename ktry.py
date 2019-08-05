#!/usr/bin/python
# -*- coding: UTF-8 -*-
from email.mime.text import MIMEText
import smtplib

def main():
    empCount = 0

def display(sss):
    sw = 'hjftest1@163.com'
    r = '1626795486@qq.com'
    p = 'qweqwe123'
    msg = MIMEText(sss, 'plain', 'utf-8')
    msg['Subject'] = 'smtp'
    msg['From'] = sw
    msg['TO'] = r
    s = smtplib.SMTP_SSL('smtp.163.com', 465)
    s.login(sw, p)
    s.sendmail(sw, r, msg.as_string())

if __name__ == '__main__':
  main()
