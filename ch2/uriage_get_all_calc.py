import openpyxl as excel

# 式の計算結果が得られるようにワークブックを読み込む
book = excel.load_workbook(
    "ch2/data/uriage.xlsx",
    data_only=True)
