import openpyxl as excel

# ワークブック（Excelファイル）を開く
book = excel.load_workbook("ch2/data/hello.xlsx")

# 先頭のワークシートを取り出す
sheet = book.worksheets[0]

# A1のセルの値を得る
cell = sheet["A1"]

# 読み出した結果を画面に出力
print(cell.value)
