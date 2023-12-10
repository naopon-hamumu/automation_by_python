import openpyxl as xl
# ブックを作りシートを得る
book = xl.Workbook()
sheet = book.active

# A1、B1、C1にすべて同じ値を設定
val = 3.14159
sheet.append([val, val, val])

# 小数点以下を省略して表示
sheet["A1"].number_format = '0'
# 小数点以下を2桁だけ表示
sheet["B1"].number_format = '0.00'
# 小数点以下を4桁だけ表示
sheet["C1"].number_format = '0.0000'

# 保存
book.save("ch2/data/number_format1.xlsx")
