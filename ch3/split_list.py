import openpyxl as excel, json
# 入出力ファイルを指定
in_file = 'ch3/data/matome.xlsx'
out_file = 'ch3/data/matome.json'

# メイン処理
def split_list():
    # Excelシートのデータを顧客ごとに分ける
    users = read_and_split(in_file)
    # お客ごとにデータを集計する
    result = {}
    for name, rows in users.items():
        result[name] = calc_user(rows)
        print(name, result[name]['total'])
    # ファイルに結果を保存
    with open(out_file, "wt") as fp:
        json.dump(result, fp)

# 入力ファイルを読んで顧客ごとに分類
def read_and_split(in_file):
    users = {} # 辞書型の変数を初期化
    # ブックを開いてシートの全行を読む
    sheet = excel.load_workbook(in_file).active
    for row in sheet.iter_rows():
        # シートの1行を取り出してリストに変換
        values = [col.value for col in row]
        name = values[1] # 顧客名を取り出す
        # 初出ならリスト初期化
        if name not in users: users[name] = []
        # データを顧客ごとに分ける
        users[name].append(values)
    return users

# 顧客1人分のデータを集計する
def calc_user(rows):
    total = 0
    items = [] # 請求の明細用
    # 集計処理を行う
    for row in rows:
        # 請求書に必要な項目だけ抽出して請求書明細の形式で追加
        date, _, item, cnt, price, _ = row
        date_s = date.strftime('%m/%d')
        items.append([date_s, item, cnt, price])
        # 合計金額を計算
        total += cnt * price
    # 集計結果を辞書形式で返す
    return {'items': items, 'total': total}

if __name__ == "__main__":
    split_list() # メイン処理を実行
