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
  print(sheet["セル名"].value)

  v = sheet.cell(row=行番号, column=列番号).value
  print(v)

  # 指定方法
  rows = sheet["左上セル名1":"右下セル名2"]
  rows = sheet["左上セル名1:右下セル名2"]
  ```

  - 指定範囲を取得する方法
    ```
    # 行番号・列番号を指定してイテレータを取得
    it = sheet.iter_rows(
      min_row=最小行, max_row=最大行,
      min_col=最小列, max_col=最大列)

    # for文と組み合わせてセルの値を得る
    for row in it:
        for cell in row:
            print(cell.value)
    ```

- セル名から行番号と列番号を取得する方法
  ```
  # セル名からセルオブジェクトを得る
  cell = sheet["セル名"]
  print(cell.row, cell.column)

  # 行番号と列番号からセルオブジェクトを得る
  cell =sheet.cell(row=行番号, column=列番号)
  print(cell.coordinate)
  ```

- Excelファイルを保存する方法
  ```
  # Excelファイルへの書き込み
  book.save("ファイル名.xlsx")
  ```
