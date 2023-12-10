### openpyxl
ExcelファイルをPythonで操作するためのライブラリ

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

#### ワークブックを扱う方法
- openpyxlを取り込む
  ```
  import openpyxl as excel
  ```

- 新規ワークブックを作成
  ```
  book = excel.Workbook()
  ```

- 既存のワークブックをファイルから開く
  ```
  book = excel.load_workbook("ファイル名.xlsx")
  ```

- ワークブックを開く（式があれば展開して開く）
  ```
  book = excel.load_workbook(
      "ファイル名.xlsx", data_only=True)
  ```

- ワークブックを明示的に閉じる
  ```
  book.close()
  ```

#### ブックからシートを選ぶ方法
- アクティブなワークシートを得る
  ```
  sheet = book.active
  ```

- 任意の箇所にあるワークシートを得る（0起点）
  ```
  sheet = book.worksheets[シート番号]
  ```

- シート名（Sheet1など）を指定して取得
  ```
  sheet = book["シート名"]
  ```

- ブック内のシート名の一覧を得る
  ```
  print(book.sheetnames)
  ```

#### 新規シート作成、コピー、削除とシート名変更
- 新規シートを作成
  ```
  sheet = book.create_sheet(title="シート名")
  ```

- 既存のシートをコピーして得る
  ```
  sheet = book.copy_worksheet(book["シート名"])
  ```

- シート名を変更する
  ```
  sheet.title = "新しい名前"
  ```

- シートを削除
  ```
  book.remove(book["シート名"])
  ```

#### セルの表示形式の設定
- セルの数値初期の設定（0.00の書式に設定する）
  ```
  sheet["セル名"].number_format = "0.00
  ```

- 日付データの書式設定（年/月/日の書式にする）
  ```
  sheet["セル名"].number_format = "yyyy/mm/dd"
  ```
