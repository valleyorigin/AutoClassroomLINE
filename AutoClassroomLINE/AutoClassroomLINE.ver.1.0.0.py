#https://di-acc2.com/system/rpa/3031/より

import imapclient
from backports import ssl
from OpenSSL import SSL
import pyzmail
import pandas as pd
import time
import pyautogui

context = ssl.SSLContext(SSL.TLSv1_2_METHOD)

imap = imapclient.IMAPClient("imap.gmail.com",ssl=True,ssl_context=context)

my_mail = "autoclassroomline@gmail.com" #メールアドレス
app_password = "jtshuwspupanhcnf"    #パスワード

imap.login(my_mail,app_password)

#print(imap.list_sub_folders()) #フォルダを表示

imap.select_folder("INBOX", readonly=True)

#① 検索キーワードを設定 & 検索キーワードに紐づくメールID検索
KWD = imap.search(["UNSEEN", "FROM", "no-reply@classroom.google.com"])
#② メールID→メール本文取得
raw_message = imap.fetch(KWD,["BODY[]"])

From_list = []
Cc_list = []
Bcc_list = []
Subject_list = []
Body_list = []

#検索結果保存
for j in range(len(KWD)):
    
    #特定メール取得
    message = pyzmail.PyzMessage.factory(raw_message[KWD[j]][b"BODY[]"])
    
    #宛先取得
    From = message.get_addresses("from")
    From_list.append(From)
    
    Cc = message.get_addresses("cc")
    Cc_list.append(Cc)
    
    Bcc = message.get_addresses("bcc")
    Bcc_list.append(Bcc)
    
    #件名取得
    Subject = message.get_subject()
    Subject_list.append(Subject)
    
    #本文
    Body = message.text_part.get_payload().decode(message.text_part.charset)
    Body_list.append(Body)
    

#④出力
table = pd.DataFrame({#"From":From_list,
                      #"Cc":Cc_list,
                      #"Bcc":Bcc_list,
                      "Subject":Subject_list,
                      "Body":Body_list,
                     })

print(table)