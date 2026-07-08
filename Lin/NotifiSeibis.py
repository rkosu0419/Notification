#################################################
# 予約通知スクリプト                              #
#                                               #
# yyyymmdd XXXXXX corp r.kosugi                 #
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
import json

##################### データ #####################
# 当日の日付
today = datetime.datetime.now()
todayStr = today.strftime("%Y%m%d")
today21 = datetime.datetime(today.year, today.month, today.day, 21, 0, 0)

# 翌日の日付
tomorrow = today + datetime.timedelta(days=1)
tomorrowStr = tomorrow.strftime('%Y%m%d')

# 昨日の日付
yesterday = today - datetime.timedelta(days=1)
yesterdayStr = yesterday.strftime('%Y%m%d')
yFileName = "./data/" + yesterdayStr + ".txt"

# SEIBIS情報取得
wb = load_workbook('Notifi.xlsx')
shtS = wb['SEIBIS']
shtC = wb['chatW']

url1 = shtS['H6'].value
url2 = shtS['H8'].value
url3 = shtS['H10'].value
Email = shtS['H12'].value
pw = shtS['H14'].value
cN = shtS['H16'].value
tPkey = shtS['H18'].value

# 辞書型リクエスト情報
payload1 = {'userId': Email,\
            'password': pw,\
            'ctrName': cN,\
            'transitionPublishKey': tPkey
            }

payload2 = {
    "innerTopPageDto": {
        "targetDate": todayStr
    }
}

payload3 ={
    "innerTopPageDto": {
        "targetDate": tomorrowStr
    }
}

# chatWork情報取得
url4 = shtC['H6'].value
apiToken = shtC['H12'].value
roomId = shtC['H14'].value
url4 = url4 + str(roomId) + "/messages"

# 予約情報抽出パターン
nonReservP = "\"innerVisitScheduleDtoList\":[],"    # 予約なし
vTimeP = "\"visitDateTime\":\"(.*?)\","             # 来店時間
crTimeP = "\"carReturnDateTime\":\"(.*?)\","        # 返却時間
mNameP = "\"memberNameKana\":\"(.*?)\","            # 氏名
orderP = "\"workTypeName\":\"(.*?)\","              # 作業
carTypeP = "\"modelType\":\"(.*?)\","               # 車種
rIdP = "\"reserveId\":\"(.*?)\","                   # 予約ID

################## 通知関数(7-21) ##################
def send_notification7to21(today, url3, url4, apiToken, nonReservP, vTimeP, crTimeP, mNameP, orderP, carTypeP, rIdP, res):
        # 予約情報抽出
    vTime = re.findall(vTimeP, res)
    crTime = re.findall(crTimeP, res)
    mName = re.findall(mNameP, res)
    mName = ",".join(mName)
    mName = mName.replace("\u3000", " ")
    mName = mName.split(",")
    order = re.findall(orderP, res)
    carType = re.findall(carTypeP, res)
    rId = re.findall(rIdP, res)
    todaySlash = today.strftime('%Y/%m/%d')

    headers = {
        "accept": "application/json",
        "content-type": "application/x-www-form-urlencoded",
        "x-chatworktoken": apiToken
    }

    # 予約なしor予約キャンセル
    if nonReservP in res:
        payload4 = {
            "self_unread": "0",
            "body": "本日(" + todaySlash + ")の予約はありません。"
        }

        requests.post(url4, data=payload4, headers=headers)
    # 予約数分の通知
    else:
        payload4 = {
            "self_unread": "0",
            "body": "本日(" + todaySlash + ")の予約更新がありました。"
        }

        requests.post(url4, data=payload4, headers=headers)

        for i in range(len(vTime)):

            # 電話番号取得リクエスト
            res3 = session.get(url3 + rId[i])
            soup = BeautifulSoup(res3.text, 'html.parser')
            # 携帯電話
            mPhone = soup.find('input', {'name':'mobileTelNumber'})['value']
            # 固定電話
            phone = soup.find('input', {'name':'telNumber'})['value']

            payload5 = {
                "self_unread": "0",
                "body": "来店時間:" + vTime[i] + "\n返却時間:" + crTime[i] + "\n氏名:" + mName[i] + "\n作業:" + order[i] + "\n車種:" + carType[i] + "\n携帯電話:" + mPhone + "\n固定電話:" + phone
            }

            requests.post(url4, data=payload5, headers=headers)

################## 通知関数(21-0) ##################
def send_notification21to0(tomorrow, url3, url4, apiToken, nonReservP, vTimeP, crTimeP, mNameP, orderP, carTypeP, rIdP, res):
        # 予約情報抽出
    vTime = re.findall(vTimeP, res)
    crTime = re.findall(crTimeP, res)
    mName = re.findall(mNameP, res)
    mName = ",".join(mName)
    mName = mName.replace("\u3000", " ")
    mName = mName.split(",")
    order = re.findall(orderP, res)
    carType = re.findall(carTypeP, res)
    rId = re.findall(rIdP, res)
    tomorrowSlash = tomorrow.strftime('%Y/%m/%d')

    headers = {
        "accept": "application/json",
        "content-type": "application/x-www-form-urlencoded",
        "x-chatworktoken": apiToken
    }

    # 予約なしor予約キャンセル
    if nonReservP in res:
        payload4 = {
            "self_unread": "0",
            "body": "明日(" + tomorrowSlash + ")の予約はありません。"
        }

        requests.post(url4, data=payload4, headers=headers)
    # 予約数分の通知
    else:
        payload4 = {
            "self_unread": "0",
            "body": "明日(" + tomorrowSlash + ")の予約更新がありました。"
        }

        requests.post(url4, data=payload4, headers=headers)

        for i in range(len(vTime)):

            # 電話番号取得リクエスト
            res3 = session.get(url3 + rId[i])
            soup = BeautifulSoup(res3.text, 'html.parser')
            # 携帯電話
            mPhone = soup.find('input', {'name':'mobileTelNumber'})['value']
            # 固定電話
            phone = soup.find('input', {'name':'telNumber'})['value']

            payload5 = {
                "self_unread": "0",
                "body": "来店時間:" + vTime[i] + "\n返却時間:" + crTime[i] + "\n氏名:" + mName[i] + "\n作業:" + order[i] + "\n車種:" + carType[i] + "\n携帯電話:" + mPhone + "\n固定電話:" + phone
            }

            requests.post(url4, data=payload5, headers=headers)

###################################################
# 昨日データファイル削除
if os.path.exists(yFileName):
    os.remove(yFileName)

############### スクレイピングスタート ##############

# セッション開始
session = requests.session()
# ホームpost
res1 = session.post(url1, data=payload1)

# 21:00以前
if today < today21:
    # 当日予約情報取得
    res2 = session.post(url2, json=payload2)
    # ファイル名
    tFileName = "./data/" + todayStr + ".txt"

    # 当日データなし：新規作成/書込み/通知
    if not os.path.isfile(tFileName):
        newData = open(tFileName, 'w', encoding='utf-8')
        newData.write(res2.text)
        newData.close()

        send_notification7to21(today, url3, url4, apiToken, nonReservP, vTimeP, crTimeP, mNameP, orderP, carTypeP, rIdP, res2.text)
        sys.exit()

    # 当日データあり：比較/更新or何もしない
    else:
        archive = open(tFileName, 'r', encoding='utf-8')
        
        # 更新なし：終了
        if archive.read() == res2.text:
            archive.close()
            sys.exit()
        # 更新あり：上書き/通知
        else:
            newData = open(tFileName, 'w', encoding='utf-8')
            newData.write(res2.text)
            newData.close()

            send_notification7to21(today, url3, url4, apiToken, nonReservP, vTimeP, crTimeP, mNameP, orderP, carTypeP, rIdP, res2.text)
            sys.exit()
# 21:00以降
elif today > today21:
    # 明日予約情報取得
    res2 = session.post(url2, json=payload3)
    # ファイル名
    tFileName = "./data/" + tomorrowStr + ".txt"

    # 明日データなし：新規作成/書込み/通知
    if not os.path.isfile(tFileName):
        newData = open(tFileName, 'w', encoding='utf-8')
        newData.write(res2.text)
        newData.close()

        send_notification21to0(tomorrow, url3, url4, apiToken, nonReservP, vTimeP, crTimeP, mNameP, orderP, carTypeP, rIdP, res2.text)
        sys.exit()

    # 当日データあり：比較/更新or何もしない
    else:
        archive = open(tFileName, 'r', encoding='utf-8')
        
        # 更新なし：終了
        if archive.read() == res2.text:
            archive.close()
            sys.exit()
        # 更新あり：上書き/通知
        else:
            newData = open(tFileName, 'w', encoding='utf-8')
            newData.write(res2.text)
            newData.close()

            send_notification21to0(tomorrow, url3, url4, apiToken, nonReservP, vTimeP, crTimeP, mNameP, orderP, carTypeP, rIdP, res2.text)
            sys.exit()
