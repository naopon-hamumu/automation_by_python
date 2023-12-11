import openpyxl as xl
# ブックを作りシートを得る
book = xl.Workbook()
sheet = book.active

# A1,B1,C1,D1に同じ日時の値を指定
import datetime
dt = datetime.datetime(
    year=2023, month=4, day=5,
    hour=11, minute=22, second=33)
sheet .append([dt, dt, dt, dt])

# 日付を「yyyy/mm/dd」で指定する
sheet["A1"].number_format = 'yyyy/mm/dd'

# 日付を「yyyy年mm月dd日」で指定する
sheet["B1"].number_format = 'yyyy年mm月dd日'

# 時間を「hh:mm:ss」で指定する
sheet["C1"].number_format = 'hh:mm:ss'

# 日付と時間を「mm/dd hh:mm:ss」で指定する
sheet["D1"].number_format = 'mm/dd hh:mm:ss'

# 保存
book.save("ch2/data/number_format_datetime.xlsx")
