import openpyxl as xl 

book = xl.Workbook()
sheet = book.active


###セルのデータタイプ

# 数値の設定
cell = sheet["A1"]
cell.value = 345
sheet["B1"] = "data_type="+cell.data_type

# 文字列を設定
cell = sheet["A2"]
cell.value = "abc"
sheet["B2"] = "data_type="+cell.data_type

# 日時を設定
cell = sheet["A3"]
from datetime import date
cell.value = date(2021,4,1)
sheet["B3"] = "data_type="+cell.data_type

book.save('data_type.xlsx')

# data_type 
# n　は　number（数値）
# s　は　string （文字列）
# d　は　datetime（日付型）