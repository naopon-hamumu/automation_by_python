### openpyxl
ExcelファイルをPythonで操作するためのライブラリ

- 使用方法
  ```
  import openpyxl as excel
  ```

- ワークブックを用意する方法
  ```
  # 新規ワークブックの作成
  book = excel.Workbook()

  # 既存のExcelファイルの読み込み
  book = excel.load_workbook("ファイル名.xlsx")
  ```

- ワークシートを操作するためのオブジェクトを取得する方法
  ```
  # アクティブなシートを得る
  sheet = book.active

  # n番目のシートを得る
  sheet = book.worksheets[n]

  # シート名を指定して取得
  sheet = book["シート名"]
  ```

- セル名を指定する方法
  ```
  # シートに値を書き込む
  sheet["セル名"] = "こんにちは" # セル名
  sheet.cell(row=行番号, column=列番号, value="こんにちは") # 行番号と列番等
  cell = sheet.cell(row=行番号, column=列番号) # 先に任意のセルを取得し、セルに値を設定する
  cell.value = "こんにちは"

  # シートの値を読む
  print(sheet["セル名"])
  ```

- Excelファイルを保存する方法
  ```
  # Excelファイルへの書き込み
  book.save("ファイル名.xlsx")
  ```
