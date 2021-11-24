import gspread
import json
#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない(固定URL)
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']#ここまでテンプレ

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('.json', scope)　#jsonファイルは同じフォルダ内に入れておく

#OAuth2の資格情報を使用してGoogle APIにログイン
gc = gspread.authorize(credentials)

#スプレッドシートキーを指定してワークブックを選択(スプレッドシートURLより)
workbook = gc.open_by_key('d/ｺｺ/edit')

# シートの一覧を一次元配列に格納する
worksheet_list = workbook.worksheets()
#ワークシート指定
worksheet = workbook.sheet1
#読み込みはこれ(セルをダイレクトに指定)
cell_value = worksheet.acell('B1').value
print(cell_value) #print忘れずに

#参考　https://tanuhack.com/library-gspread/#i-9 (2021.11.25)
