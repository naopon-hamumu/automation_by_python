import openpyxl as excel

# ワークブックを開いてシートを取り出す
book = excel.load_workbook("ch2/data/test100.xlsx")
sheet = book.active

# イテレータを取得する
it = sheet.iter_rows(
        min_row=2, min_col=2,
        max_row=4, max_col=4)

# for文で繰り返し値を得る
for row in it:
    r = []
    for cell in row:
        r.append(cell.value)
    print(r)
