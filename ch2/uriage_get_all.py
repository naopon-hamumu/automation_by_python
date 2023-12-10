import openpyxl as excel

# 売り上げデータのブックを開いてシートを取り出す
book = excel.load_workbook("ch2/data/uriage.xlsx")
sheet = book.active

# A3からF9のセルを取り出す
rows = sheet["A3":"F9"]
for row in rows:
    # セルの値をリストとして得る
    values = [cell.value for cell in row]
    # リストを表示する
    print(values)
