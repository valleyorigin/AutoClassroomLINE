#送信メッセージを変更

import imapclient
from backports import ssl
from OpenSSL import SSL
from urlextract import URLExtract
import pyzmail
import time
import pyautogui
import pandas as pd
import requests

def check_mail():
    global imap
    
    title = ""

    #解析メールの結果保存用
    From_list = []
    Cc_list = []
    Bcc_list = []
    Subject_list = []
    Body_list = []
    context = ssl.SSLContext(SSL.TLSv1_2_METHOD)

    imap = imapclient.IMAPClient("imap.gmail.com",ssl=True,ssl_context=context)

    my_mail = "クラスルームの通知のメールを受信するメールアドレス@gmail.com" #メールアドレス
    app_password = "上記のメールアドレスがgmailの場合、アプリパスワードを使用。そのアプリパスワード（アルファベット）をここに入力"    #パスワード

    imap.login(my_mail,app_password)

    imap.select_folder("INBOX",readonly=False)   #これがないと、動かない。

    KWD = imap.search(["UNSEEN","FROM","no-reply@classroom.google.com"])
    raw_message = imap.fetch(KWD,["BODY[]"])

    #検索結果保存
    for j in range(len(KWD)):
        
       #特定メール取得
        message = pyzmail.PyzMessage.factory(raw_message[KWD[j]][b"BODY[]"])
        
       #件名取得
        Subject = message.get_subject()
        Subject_list.append(Subject)
        #本文
        Body = message.text_part.get_payload().decode(message.text_part.charset)
        Body_list.append(Body)
    

    for title in Body_list:
        print("<メールを受け取りました。>")

    return title

title = ""
TOKEN = "LINE Notifyのトークンをここに入力"   #トークン
api_url = "https://notify-api.line.me/api/notify"
count = 0

print("自動でGoogleClassroomの投稿をLINEに投稿するプログラムです。")

while True:
    title = check_mail()
    if title == "":
        title = ""
    else:
        print(f"{title}\r\n↑メール本文")
        
        target = '新しいお知らせ\r\n'   #本文をトリミング
        idx = title.find(target)
        Send_Text = title[idx+len(target):]
        Send_Text = Send_Text[:Send_Text.find("\r\n開く")]

        target = 'さん\r\n' #先生の名前を抽出
        idx = title.find(target)
        Send_title = title[idx+len(target):]
        Send_title = Send_title[:Send_title.find("先生が")]
        
        
        #URLを取得
        extractor = URLExtract()
        extractor.extract_localhost = False
        ALL_URL = extractor.find_urls(title)
        Send_URL = ALL_URL[-2]

        message =  f"{Send_title}先生がお知らせを投稿しました。\r\n{Send_Text}\r\nURL\r\n{Send_URL}"
        #message = Send_URL

        print(message)
        #LINEで送信するプログラム
        
        send_contents = message   #通知内容
        TOKEN_dic = {"Authorization": "Bearer" + " " + TOKEN}
        send_dic = {"message": send_contents}
        requests.post(api_url, headers=TOKEN_dic, data=send_dic)
        count = count + 1
        title = ""
                

print("メッセージ送信回数 : ", count)

#print(title)
#titleは本文
