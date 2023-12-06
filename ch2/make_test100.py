import openpyxl as excel

book = excel.Workbook()
sheet = book.active

# 連続でセルに値を設定する
for y in range(1, 101):
    for x in range(1, 101):
        cell = sheet.cell(row=y, column=x)
        cell.value = cell.coordinate # セル名

# ファイルを保存
book.save("ch2/data/test100.xlsx")
