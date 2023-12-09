import openpyxl as excel

# ワークブックを開いてシートを得る
book = excel.load_workbook("ch2/data/test100.xlsx")
sheet = book.active

# 連続でセルの値を得る
for row in sheet["B2":"D4"]:
    print([c.value for c in row])
