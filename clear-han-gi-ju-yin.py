import xlwings as xw 

# 開啟 Excel 檔案
# connect to a file that is open or in the current working directory
#wb = xw.Book()  # create a new workbook
wb = xw.Book('漢文注音工具.xlsx')  

#################################################
# 清除漢文
#################################################

# 指定「作用工作表」
# Instantiate a sheet object
# sheet = wb.sheets['工作表1']
sheet = wb.sheets['文章']

# 針對 Range ，清除各儲存格的內容值
# Reading/writing values from/to ranges 
# sheet.range('A1').value = '軟體開發專案'
sheet.range('D4').value = ''

#################################################
# 清除漢文注音
#################################################

# 指定「作用工作表」
# Instantiate a sheet object
# sheet = wb.sheets['工作表1']
sheet = wb.sheets['【拼音】注音文章']

# 針對 Range ，清除各儲存格的內容值
# Reading/writing values from/to ranges 
# sheet.range('A1').value = '軟體開發專案'
sheet.range('D9:F1588').value = ''
