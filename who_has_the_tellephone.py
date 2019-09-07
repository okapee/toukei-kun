import configparser
import codecs
import pandas as pd
import pymsteams
import datetime
import schedule

# config.ini読込み
inifile = configparser.ConfigParser()
# inifile.read('./config.ini')
inifile.readfp(codecs.open("config.ini", "r", "utf8"))

# def who_has_phone():
# 一般チャネルへ投稿
webhook = inifile.get('settings', 'post_common')

# Excelファイルのロード(読み取り専用)
excel_path = inifile.get('settings', 'excel_path')
df = pd.read_excel(excel_path, index_col=0)
print(df)

# # 行ラベル値
# print(df.index.values)
# # 列ラベル値
# print(df.columns.values)

# 日付の取得
today = datetime.datetime.today()
# print(today.day)

sei_name = str(df[df[today.day] == '保正'].index[0])
fuku_name = str(df[df[today.day] == '保副'].index[0])

# 当日の保守携帯の正／副当番をチャネルに投稿する
myTeamsMessage = pymsteams.connectorcard(webhook)
myTeamsMessage.title("本日の保守携帯当番")

# disignate HEX Color Code
myTeamsMessage.color("#7fb9b9")

# 保守携帯保持者表示
myTeamsMessage.text("保守携帯正： " + sei_name + "、保守携帯副： " + fuku_name)

# リンクボタン
myTeamsMessage.addLinkButton("保守携帯当番表はこちら", "https://ts.accenture.com/:x:/r/sites/JPS_Statistical_Management_System/2000Delivery/50_%E7%B5%B1%E8%A8%88AP%E4%BF%9D%E5%AE%88/30_%E7%AE%A1%E7%90%86/%E3%80%90Legatus%E3%80%91%E4%BC%91%E6%9A%87%E7%AE%A1%E7%90%86%E5%85%BC%E4%BF%9D%E5%AE%88%E6%90%BA%E5%B8%AF%E5%BD%93%E7%95%AA%E3%82%AB%E3%83%AC%E3%83%B3%E3%83%80%E3%83%BC.xlsx?d=w3c5504bd00724e4586bf2232c21c2154&csf=1")

myTeamsMessage.printme()
myTeamsMessage.send()

# Cron化
# schedule.every().day.at("12:00").do(who_has_phone)