import openpyxl as xl 

book = xl.Workbook()
sheet = book.active

import datetime
dt = datetime.datetime(
    year = 2023,month = 4,day = 5,
    hour = 11,minute = 22, second = 33 
)
sheet.append([dt,dt,dt,dt])

################### 日付を指定  ###############

sheet["A1"].number_format = "yyyy/mm/dd"

sheet["B1"].number_format = "yyyy年mm月dd日"

sheet["C1"].number_format = "hh:mm:ss"

sheet["D1"].number_format = "mm/dd hh:mm:ss"

book.save('number_format_dateatime.xlsx')