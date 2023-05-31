import openpyxl as xl 

book = xl.Workbook()
sheet = book.active

# A1,B1,C1にすべてを設定
val = 3.14159
sheet.append([val,val,val])

# number_formatは小数点以下の表示を指定できる　指定方法は”0.0000”で指定する

sheet["A1"].number_format = "0"         # 小数点以下を省略して表示

sheet["B1"].number_format = "0.00"      # 小数点以下を２桁だけ表示

sheet["C1"].number_format = "0.0000"    # 小数点以下を４桁だけ表示

book.save('number_format.xlsx')