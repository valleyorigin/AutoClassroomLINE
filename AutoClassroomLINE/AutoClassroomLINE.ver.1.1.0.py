#件名を取得する

import imapclient
from backports import ssl
from OpenSSL import SSL
import time
import pyautogui
import pyzmail
import pandas as pd

def check_mail():
    title = ""

    #解析メールの結果保存用
    From_list = []
    Cc_list = []
    Bcc_list = []
    Subject_list = []
    Body_list = []

    context = ssl.SSLContext(SSL.TLSv1_2_METHOD)

    imap = imapclient.IMAPClient("imap.gmail.com",ssl=True,ssl_context=context)

    my_mail = "autoclassroomline@gmail.com" #メールアドレス
    app_password = "jtshuwspupanhcnf"    #パスワード

    imap.login(my_mail,app_password)

    #print(imap.list_sub_folders()) #フォルダを表示

    imap.select_folder("INBOX",readonly=True)   #これがないと、動かない。

    KWD = imap.search(["UNSEEN","FROM","no-reply@classroom.google.com"])
    raw_message = imap.fetch(KWD,["BODY[]"])

    #検索結果保存
    for j in range(len(KWD)):
        
       #特定メール取得
        message = pyzmail.PyzMessage.factory(raw_message[KWD[j]][b"BODY[]"])
    
        #宛先取得
        #From = message.get_addresses("from")
        #From_list.append(From)
    
        #Cc = message.get_addresses("cc")
        #Cc_list.append(Cc)
    
        #Bcc = message.get_addresses("bcc")
        #Bcc_list.append(Bcc)
    
       #件名取得
        Subject = message.get_subject()
        Subject_list.append(Subject)
    
        #本文
        #Body = message.text_part.get_payload().decode(message.text_part.charset)
        #Body_list.append(Body)

    #print(Subject_list)

    for title in Subject_list:
        True

    #print(title)    #titleをLINEで送信できるようにすればいい。
    
    return title

check_mail()

print("関数を実行した結果")
title = check_mail()
print(title)
#titleは本文