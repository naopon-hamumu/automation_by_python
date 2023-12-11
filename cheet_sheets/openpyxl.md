### openpyxl
ExcelファイルをPythonで操作するためのライブラリ

#### マニュアル
  - [openpyxlのマニュアル](https://openpyxl.readthedocs.io/)
  - [openpyxl内のモジュール一覧](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.html)

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

  | 記号 | 意味 |
  | :---: | :---: |
  | yyyy | 西暦(4桁) |
  | mm | 月(2桁) |
  | m | 月(1桁) |
  | dd | 日(2桁) |
  | d | 日(1桁) |
  | dddd | 曜日(英語) |
  | mmm | 英語の月名 |
  | [$-411]ggge | 和暦年 |
  | [$-411]gge | 和暦年(1文字) |
  | [$411]dddd | 曜日(日本語) |

- 罫線の書式設定<br>
  | 普通の線 | thick, thin, medium. double |
  | 点線 | dashed, dotted, mediumDashDot, mediumDashDotDot, slantDashDot, mediumDashed, dashDotDot, dashDot |
