#################################################
# XXXX 予約通知スクリプト                        #
#                                               #
# yyyymmdd xxxxxx corp r.kosugi                 #
#################################################

################### 日本語対応 ###################
#!/usr/bin/env python
# -*- coding: utf-8 -*-

##################### import ####################
import os
import sys
import requests
from bs4 import BeautifulSoup
import re
import datetime
from openpyxl import load_workbook
import csv

##################### データ #####################
# 当日の日付
today = datetime.datetime.now()
todayStr = today.strftime("%Y%m%d")
today21 = datetime.datetime(today.year, today.month, today.day, 21, 0, 0)
tFileName = "./data/" + todayStr + ".csv"
todaySlash = today.strftime('%Y/%m/%d')

# 昨日の日付
yesterday = today - datetime.timedelta(days=1)
yesterdayStr = yesterday.strftime('%Y%m%d')
yFileName = "./data/" + yesterdayStr + ".csv"

# NICORAS情報取得
wb = load_workbook('Notifi.xlsx')
shtN = wb['NICO']
shtC = wb['chatW']

url1 = shtN['H6'].value
url2 = shtN['H8'].value
page_id = shtN['H16'].value
actionCount = shtN['H18'].value
login_stage = shtN['H20'].value
user_id = shtN['H12'].value
user_pass = shtN['H14'].value
recaptcha = shtN['H22'].value
login_btn = shtN['H24'].value

# chatWork情報取得
url4 = shtC['H6'].value
apiToken = shtC['H12'].value
roomId = shtC['H16'].value
url4 = url4 + str(roomId) + "/messages"

# 辞書型リクエスト情報
payload1 = {'page_id': page_id,\
            'actionCount': actionCount,\
            'login_stage': login_stage,\
            'user_id': user_id,\
            'user_pass': user_pass,\
            'g-recaptcha-response': recaptcha,\
            'login_btn': login_btn
            }

headers = {
    "accept": "application/json",
    "content-type": "application/x-www-form-urlencoded",
    "x-chatworktoken": apiToken
}


# 予約情報抽出パターン
timeP = "<a href=\"admin_book_info.php\" .*?>(.*?)          <br />"    # 予約時間
nameP = "-->(.*?)</a>"                                                 # お名前
carTypeP = r"\(\'go_admin_book_info\',(.*?)\);"                        # 車両

############### carTyeNum車番変換関数 ##############
def carNumConv(carTypeNum):   
    for i in range(len(carTypeNum)):
        if re.match("\'120\',\'0\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "K N-BOX 八王子 580 わ 3418"
        elif re.match("\'120\',\'1\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "K N-BOX 八王子 580 わ 4019"
        elif re.match("\'120\',\'2\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "K N-BOX 八王子 580 わ 4018"
        elif re.match("\'120\',\'3\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "K N-BOX 八王子 580 わ 4017"
        elif re.match("\'120\',\'4\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "K スペーシア 八王子 580 わ 4211"
        elif re.match("\'120\',\'5\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "K スペーシア 八王子 580 わ 4210"
        elif re.match("\'100\',\'0\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ヤリス 八王子 501 わ 622"
        elif re.match("\'100\',\'1\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S アクア 八王子 501 わ 1458"
        elif re.match("\'100\',\'2\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S アクア 八王子 501 わ 1459"
        elif re.match("\'100\',\'3\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S アクア 八王子 501 わ 1460"
        elif re.match("\'100\',\'4\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S アクア 八王子 501 わ 2183"
        elif re.match("\'100\',\'5\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ライズ 八王子 501 わ 2489"
        elif re.match("\'100\',\'6\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ライズ 八王子 501 わ 2488"
        elif re.match("\'100\',\'7\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ライズ 八王子 501 わ 2490"
        elif re.match("\'100\',\'8\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ライズ 八王子 501 わ 2491"
        elif re.match("\'100\',\'9\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ライズ 八王子 501 わ 2492"
        elif re.match("\'100\',\'10\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ソリオハイブリッド 八王子 501 わ 2854"
        elif re.match("\'100\',\'11\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ソリオハイブリッド 八王子 501 わ 2855"
        elif re.match("\'100\',\'12\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ソリオハイブリッド 八王子 501 わ 2856"
        elif re.match("\'100\',\'13\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ソリオハイブリッド 八王子 501 わ 2857"
        elif re.match("\'100\',\'14\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "S ソリオハイブリッド 八王子 501 わ 2858"
        elif re.match("\'410\',\'0\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "F1 ノア 八王子 300 わ 3563"
        elif re.match("\'410\',\'1\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "F1 ヴォクシー 八王子 300 わ 4087"
        elif re.match("\'410\',\'2\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "F1 ノア 八王子 501 わ 2385"
        elif re.match("\'510\',\'0\',\'.*\'",carTypeNum[i]):
            carTypeNum[i] = "T1 (軽トラック)キャリイ 八王子 480 わ 1832"
    return carTypeNum

##################### 通知関数 #####################
def send_notification(todaySlash, time, name, carType):
    # 予約なし
    if not time:
        payload2 = {
            "self_unread": "0",
            "body": "本日(" + todaySlash + ")の予約はありません。"
        }
        # 通知送信
        requests.post(url4, data=payload2, headers=headers)
    # 予約あり    
    else:
        payload2 = {
            "self_unread": "0",
            "body": "本日(" + todaySlash + ")の予約内容です。"
        }
        requests.post(url4, data=payload2, headers=headers)

        # 予約数通知
        for i in range(len(time)):
            payload2 = {
                "self_unread": "0",
                "body": "予約時間：" + time[i] + "\nお名前：" + name[i] + "\n予約車両：" + carType[i]
            }
            requests.post(url4, data=payload2, headers=headers)

###################################################
# 昨日データファイル削除
if os.path.exists(yFileName):
    os.remove(yFileName)

############### スクレイピングスタート ##############
# セッション開始
session = requests.session()
# ログイン
res1 = session.post(url1, data=payload1)
# 日別予約状況
res2 = session.get(url2)
session.close()

# htmlから「予約時間」「お名前」「車両」を取得、二次元配列
time = re.findall(timeP, res2.text)
name = re.findall(nameP, res2.text)
carTypeNum = re.findall(carTypeP, res2.text)
# carTypeNumを車種車番へ変換
carType = carNumConv(carTypeNum)

# csv出力二次元データ
newData = [time, name, carType]

# 当日データなし(当日初)
if not os.path.isfile(tFileName):
    # csv書込み
    with open(tFileName, "w", encoding="utf_8", newline="") as nData:
        writer = csv.writer(nData)
        writer.writerow(newData)
        nData.close
    # 通知
    send_notification(todaySlash, time, name, carType)
    sys.exit()
# 当日データあり
else:
    # 前データの読み込み/リストへ加工
    with open(tFileName, "r", encoding="utf_8", newline="") as aData:
        #reader = csv.reader(aData, delimiter=",", doublequote=True, lineterminator="\r\n", quotechar='"', skipinitialspace=True)
        reader = csv.reader(aData)
        archive = [row for row in reader]
        archiveTime = archive[0][0].replace("\"", "").replace("\'", "").replace("[", "").replace("]", "").split(', ')
        archiveName = archive[0][1].replace("\"", "").replace("\'", "").replace("[", "").replace("]", "").split(', ')
        archiveCarTypr = archive[0][2].replace("\"", "").replace("\'", "").replace("[", "").replace("]", "").split(', ')
        # 再度まとめる
        reArchive = [archiveTime, archiveName, archiveCarTypr]
    aData.close

    # データ比較：更新なし
    if reArchive == newData:
        sys.exit()
    # データ比較：更新あり
    else:
        # csv上書き
        with open(tFileName, "w", encoding="utf_8", newline="") as nData:
            writer = csv.writer(nData)
            writer.writerow(newData)
            nData.close
        # 通知
        send_notification(todaySlash, time, name, carType)
        sys.exit()
