# -*- coding: cp936 -*-
import smtplib
from email.mime.text import MIMEText
from itertools import islice
from openpyxl import load_workbook
from tkinter import *
import tkinter.messagebox
import win32ui
from tkinter import *
import time, easygui
##################��������������Ȩ��######################
root = Tk()
root.title("����������")
root.geometry('300x300')  #��x ����*
l1 = Label(root, text="����")
l1.pack()  #�����side���Ը�ֵΪLEFT ?RTGHT TOP ?BOTTOM
xls_text = StringVar()
xls = Entry(root, textvariable = xls_text)
xls_text.set("")
xls.pack()
l2 = Label(root, text="��½��Ȩ��")
l2.pack()  #�����side���Ը�ֵΪLEFT ?RTGHT TOP ?BOTTOM
sheet_text = StringVar()
sheet = Entry(root, textvariable = sheet_text)
sheet_text.set("")
sheet.pack()
#sender = '943811605@qq.com'   #����������
#passwd = 'owspejljkrkrbdef'   #������������Ȩ��(��Ҫ����������������)
#mreceivers = '1626795486@qq.com'
#sender = '2861825250@qq.com'   #����������
#passwd = 'vrqusnisgesbdehc'   #������������Ȩ��(��Ҫ����������������)
#sender = 'hjftest1@126.com'   #����������
#passwd = '123123123q'   #������������Ȩ��(��Ҫ����������������)
#cc = 'hjftest1@126.com'
#sender = 'hjftest1@163.com'   #����������
#passwd = 'qweqwe123'   #������������Ȩ��(��Ҫ����������������)
#cc = 'hjftest1@163.com'

def on_click():
    #global sender
    #global passwd
    #sender = xls_text.get()
    #passwd = sheet_text.get()
    string = str("���䣺%s ��½��Ȩ�룺%s " %(sender, passwd))
    print("���䣺%s ��½��Ȩ�룺%s " %(sender, passwd))
    tkinter.messagebox.showinfo(title='aaa', message = string)
    root.destroy()
Button(root, text="�ύ", command = on_click).pack()
root.mainloop()
####################################################

####################���ʵ������������########################
tkinter.messagebox.askokcancel("ѡ���ļ�","��ѡ���ʵ�(xlsx��ʽ)")
dlg = win32ui.CreateFileDialog(1)  # 1��ʾ���ļ��Ի���
dlg.SetOFNInitialDir('C:/')  # ���ô��ļ��Ի����еĳ�ʼ��ʾĿ¼
dlg.DoModal()
filename1 = dlg.GetPathName()  # ��ȡѡ����ļ�����

wb = load_workbook(filename1)
sheet = wb.get_sheet_by_name("Sheet1")
t1_arr=[]
t2_arr=[]
for i in sheet["1"]:
    t1_arr.append(str(i.value))
for j in sheet["2"]:
    t2_arr.append(str(j.value))

sss= ''
dict ={}
for line in range(3,sheet.max_row+1):
    arr = []
    res = ''
    #print(sheet[str(line)])
    for i in sheet[line]:
        arr.append(str(i.value))
    strs = '\t'.join(arr)
    sss = sss + strs + '\n'
    arr_mess = arr[:-1]
    #print(strs)
    #print('\n')

    for n in range(0,len(arr_mess)):
        if n < len(t1_arr)-1:
            if t1_arr[n] != "None" and t1_arr[n+1] == "None":
                res = res + '--------------' + t1_arr[n]+'---------------: \n'
                res = res + '---'+t2_arr[n]+': '+arr[n] + '\n'
            elif t1_arr[n] != "None" and t1_arr[n+1] != "None":
                res = res + t1_arr[n] + ': '+arr[n] + '\n'
            elif t1_arr[n] == "None" and t1_arr[n+1] == "None":
                res = res + '---'+t2_arr[n]+': '+arr[n] + '\n'
            elif t1_arr[n] == "None" and t1_arr[n+1] != "None":
                res = res + '---' + t2_arr[n] + ': ' + arr[n] + '\n-------------------------------------\n'
        elif n == len(arr_mess)-1:
            res = res + t1_arr[n] + ': ' + arr[n] + '\n'
    dict[arr[-1]] = res
#print (dict)

#sender = '943811605@qq.com'   #����������
#passwd = 'owspejljkrkrbdef'   #������������Ȩ��(��Ҫ����������������)


sucess_num = 0
fail_num = 0
total_num = 0
fail_arr = []

'''
msg = MIMEText(sss,'plain','utf-8')
msg['Subject'] = 'smtp'
msg['From'] = msender
msg['TO'] = mreceivers
s = smtplib.SMTP_SSL('smtp.qq.com', 465)
s.login(msender,mpasswd)
s.sendmail(msender,mreceivers,msg.as_string())
'''

time1 = 0
cnum = 0
for email,mess in dict.items():
    cnum = cnum +1
    if cnum >14:
        time.sleep(160)
        cnum = 0
    else:
        time.sleep(3)
    total_num = total_num + 1
    receivers = email #�ռ�������
    print(receivers)
    subject = '���ʵ�' #����
    content = mess
    #print(content)
    msg = MIMEText(content,'plain','utf-8')
    msg['Subject'] = subject
    msg['From'] = sender
    #msg['Cc'] = cc
    msg['TO'] = receivers
    #print(msg)
    if 'zjsos' in sender:
        s = smtplib.SMTP_SSL('smtp.exmail.qq.com', 465)
    elif 'qq.com' in sender:
        s = smtplib.SMTP_SSL('smtp.qq.com', 465)
    elif '126.com' in sender:
        s = smtplib.SMTP_SSL('smtp.126.com', 465)
    elif '163.com' in sender:
        s = smtplib.SMTP_SSL('smtp.163.com', 465)
    elif '2980.com' in sender:
        s = smtplib.SMTP_SSL('smtp.2980.com', 465)
    #elif ''
    else:
        s = smtplib.SMTP_SSL('smtp.exmail.qq.com', 465)
        easygui.msgbox("��δ֧�ִ����䣬����ϵ������", title="����", ok_button="ȷ��")
    try:
        s.login(sender,passwd)
        s.sendmail(sender,receivers,msg.as_string())
        print('���ͳɹ�',sender,receivers)
        sucess_num = sucess_num+1
    except:
        print('����ʧ��')
        fail_num = fail_num +1
        fail_arr.append(email)
print("ȫ��Ա��:%s ���ͳɹ�:%s ����ʧ��:%s" %(total_num,sucess_num,fail_num))
easygui.msgbox("ȫ��Ա��:%s ���ͳɹ�:%s ����ʧ��:%s" %(total_num,sucess_num,fail_num), title="���ͽ��",ok_button="ȷ��")
if fail_num>0:
    print("����Ŀ�����䷢��ʧ�ܣ�")
    kk = '\n'.join(fail_arr)
    easygui.msgbox(kk, title="ʧ����������", ok_button="ȷ��")