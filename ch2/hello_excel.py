# ライブラリを取り込む
import openpyxl as excel

# 新規ワークブックを作る
book = excel.Workbook()

# アクティブなワークシートを得る
sheet = book.active

# A1のセルに値を設定
sheet["A1"] = "こんにちは"

# ファイルを保存
book.save("ch2/data/hello.xlsx")
