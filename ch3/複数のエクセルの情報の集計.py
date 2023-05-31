import glob
import openpyxl as excel

# # ファイルの列挙
# files = glob.glob("ch3/xlsx/*.xlsx")
# # 列挙したファイルを表示
# print(files)# 実行結果　['ch3/xlsx\\invoice-template.xlsx', 'ch3/xlsx\\invoice01.xlsx', 'ch3/xlsx\\uriage.xlsx']

# 読み込むファイル　と　セーブ先ファイル
target_dir = "ch3/salesbooks"
save_file = "ch3/xlsx/matome.xlsx"

def read_files():   #　メイン処理
    ## 売上一覧を書き込むブックを用意する
    book = excel.Workbook()
    main_sheet = book.active
    # ファイルを列挙して読む
    enumfiles(main_sheet)
    book.save(save_file)


def enumfiles(main_sheet):  #　ファイルの列挙
    files = glob.glob(f"{target_dir}/*.xlsx")

    for fname in files: # 各エクセルを読み込む
        read_book(main_sheet,fname)


def read_book(main_sheet,fname):    # ブックを開いて中身を読む
    print("read:",fname)

    book = excel.load_workbook(fname,data_only=True)
    sheet = book.active

    rows = sheet["A4":"F999"]   # 売上データのある範囲を読み込む
    for row in rows:
        values = [cell.value for cell in row]   # セルの値をリストとして得る
        if values[0] is None: break
        print(values)

        main_sheet.append(values)  #　main_sheetに値をコピー

if __name__ == "__main__":  # メインプログラムを実行　
    read_files()
    print(__name__) #ここでは__main__と表示される