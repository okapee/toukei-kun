import pandas as pd
import pymsteams
import datetime
import schedule

# def who_has_phone():
# 統計一般
webhook = "https://outlook.office.com/webhook/cade7442-15b4-4935-8113-7b42106d541b@e0793d39-0939-496d-b129-198edd916feb/IncomingWebhook/65f51b981f6c4752914e1b708e7f6eb1/09a4ff80-ffd2-4222-9aea-ee9fbe9b9466"
# 統計テスト
# webhook = "https://outlook.office.com/webhook/cade7442-15b4-4935-8113-7b42106d541b@e0793d39-0939-496d-b129-198edd916feb/IncomingWebhook/bd45252cee9347659fa3ddc8e76921a9/09a4ff80-ffd2-4222-9aea-ee9fbe9b9466"

# Excelファイルのロード(読み取り専用)
excel_path = r"C:\Users\masahiro.okazaki\Dropbox\とうけいくん\test.xlsx"
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