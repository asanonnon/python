import openpyxl as excel,json

# 入出力ファイル
in_file = "ch3/xlsx/matome.xlsx"
out_file = "ch3/xlsx/matome.json"

def split_list():   #　メイン処理
    
    users = read_and_split(in_file) #　エクセルのデータを顧客ごとに分ける
    
    # 顧客ごとにデータを集計する
    result= {}
    for name,rows in users.items():
        result[name] = calc_user(rows)  # 顧客別集計の関数に移動　返り値を代入
        print(name,result[name]["total"],result[name]["items"])

    with open(out_file,"wt") as fp: # 代入結果をファイルに書き込む
        json.dump(result,fp)



def read_and_split(in_file):            # 入力ファイルを読んで顧客ごとに分割
    users = {}                          # 辞書型の変数を初期化

    sheet = excel.load_workbook(in_file).active # ファイル読み込み
    for row in sheet.iter_rows():
        values = [col.value for col in row]     # 読んだファイルをリスト化
        name = values[1]                        #リストの顧客名を代入

        if name not in users:users[name] = []   #ない場合,顧客名の空のリストを作成
        users[name].append(values)              #顧客の連想配列に名前が一致するものに追加

    return users                                # 顧客別に分割した配列を返す


def calc_user(rows):        # 顧客一人分のデータを集計
    total = 0       #　初期化
    items = []

    for row in rows:        # 請求書にn必要な項目だけ抽出して請求書明細の形式で追加
        date, _, item, cnt, price, _= row   # データ仕分け　_のところは使わない値
        date_s = date.strftime('%m/%d') #　日時を間に/を入れて整形
        items.append([date_s, item, cnt, price])    # itemsのリストに２次元配列として追加

        total += cnt * price    #　totalに足していく
    
    return {"items": items,"total":total}   # 連想配列　itemsは日時、品名、個数、値段　totalは合計金額　を返す


if __name__ == "__main__":
    split_list()
