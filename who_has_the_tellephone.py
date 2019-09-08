import configparser
import codecs
import pandas as pd
import pymsteams
import datetime
import schedule
from selenium import webdriver
import requests
import json
import datetime

date = datetime.datetime.today().strftime("%Y/%m/%d")

# http://api.jugemkey.jp/api/horoscope/year/month/day の形式
res = requests.get(url='http://api.jugemkey.jp/api/horoscope/free/'+ date)

print(json.dumps(json.loads(res.text), indent=4, ensure_ascii=False))

# horoscope_list = ["牡羊座", "牡牛座", "双子座", "蟹座", "獅子座", "乙女座", "天秤座", "蠍座", "射手座", "山羊座", "水瓶座"]
# today_top_horoscrope = min(res.json()["horoscope"][date][0]["rank"], res.json()["horoscope"][date][1]["rank"], res.json()["horoscope"][date][2]["rank"], res.json()["horoscope"][date][3]["rank"], res.json()["horoscope"][date][4]["rank"], res.json()["horoscope"][date][5]["rank"], res.json()["horoscope"][date][6]["rank"], res.json()["horoscope"][date][7]["rank"], res.json()["horoscope"][date][8]["rank"], res.json()["horoscope"][date][9]["rank"])
# print("今日の運勢一位は、" + today_top_horoscrope + "です。")

counter = 0
top_horoscope = ""
top_color = ""
top_item = ""
while counter < 12:
    print("debug: " + str(res.json()["horoscope"][date][counter]["rank"]))
    if res.json()["horoscope"][date][counter]["rank"] == 1:
        print("debug: " + "LOOP IN!")
        top_horoscope = res.json()["horoscope"][date][counter]["sign"]
        top_color = res.json()["horoscope"][date][counter]["color"]
        top_item = res.json()["horoscope"][date][counter]["item"]
        top_content = res.json()["horoscope"][date][counter]["content"]
    counter += 1

print("本日のTOP1は" + top_horoscope + " です。")

#たとえば、牡羊座のみ取得したい場合
# print(res.json()["horoscope"][date][0]["rank"])


# #1)WebDriverを開く
# driver = webdriver.Chrome('C:\\Users\\masahiro.okazaki\\Documents\\programming\\chromedriver.exe')

# #2)取得したいサイトのurlを指定
# url = 'https://ts.accenture.com/sites/JPS_Statistical_Management_System/2000Delivery/Forms/AllItems.aspx?RootFolder=%2Fsites%2FJPS%5FStatistical%5FManagement%5FSystem%2F2000Delivery%2F50%5F%E7%B5%B1%E8%A8%88AP%E4%BF%9D%E5%AE%88%2F30%5F%E7%AE%A1%E7%90%86'
# #3)取得したいサイトを開く。Chromeのウィンドウが開く
# driver.get(url)

# #4)要素を指定して、Contentオブジェクトに代入
# Content = driver.find_element_by_xpath(
#     '//*[@id="DeltaPlaceHolderPageTitleInTitleArea"]/span/span[3]/a')

# #5)contentで取れた内容を確認
# Content.text


# config.ini読込み
inifile = configparser.ConfigParser()
# inifile.read('./config.ini')
inifile.readfp(codecs.open("config.ini", "r", "utf8"))

# def who_has_phone():
# 一般チャネルへ投稿
# webhook = inifile.get('settings', 'post_common')
# Testチャネルへの投稿
webhook = inifile.get('settings', 'post_common')

# Excelファイルのロード(読み取り専用)
excel_path = inifile.get('settings', 'excel_path')
df = pd.read_excel(excel_path, index_col=0)
print(df)

# spsloc_messageのロード
sps_loc = inifile.get('message', 'sps_loc')

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
myTeamsMessage.addLinkButton("保守携帯当番表はこちら", sps_loc)

myTeamsMessage.printme()
myTeamsMessage.send()


# 当日の保守携帯の正／副当番をチャネルに投稿する
myTeamsMessage = pymsteams.connectorcard(webhook)
myTeamsMessage.title("本日の運勢")

# disignate HEX Color Code
myTeamsMessage.color("#7fb9b9")

# 保守携帯保持者表示
myTeamsMessage.text("本日の運勢１位は… " + top_horoscope + " です！\n\n" + "ラッキーカラーは*" + top_color + "*、\n\n" + "ラッキーアイテムは*" + top_item + "*です。\n\n\n"+ top_content)


myTeamsMessage.printme()
myTeamsMessage.send()
# Cron化
# schedule.every().day.at("12:00").do(who_has_phone)
