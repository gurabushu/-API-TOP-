import requests
import pandas as pd
import os
import subprocess

#楽天API
API_KEY = ""
CATEGORY_ID = ""
API_URL = f"https://app.rakuten.co.jp/services/api/IchibaItem/Ranking/20170628?format=json&applicationId={API_KEY}&genreId={CATEGORY_ID}"


#APIリクエストを送信
response = requests.get(API_URL)

#データ取得
if response.status_code==200:
        data = response.json()

#リスト作成
products =[]


#商品情報取得
   # **リスト `Items` を正しくループ処理**
for item in data["Items"]:
        product = item["Item"]  # 商品データの辞書を取得
        rank = product["rank"]  # ランキング順位
        name = product["itemName"]  # 商品名
        price = product["itemPrice"]  # 価格
        review_count = product["reviewCount"]  # レビュー数
        review_avg = product["reviewAverage"]  # 平均評価

        products.append((rank, name, price, review_count, review_avg))


df = pd.DataFrame(products,columns=["順位","商品名","価格","レビュー数","平均評価"])

#Excel指定
new_excel = "試用スクレイピング.xlsx"

#Excelに保存
df.to_excel(new_excel)



subprocess.Popen(["start", new_excel], shell=True) 
