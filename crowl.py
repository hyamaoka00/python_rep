# coding: UTF-8
import os
import openpyxl
import edit_file

# エクセル操作
os.chdir("D:\OneDrive\sbi")
wb = openpyxl.load_workbook("investment.xlsx")
sheet = wb["買い付け候補"]

# 銘柄取得列
j = 1
# セル取得開始位置
i = 2
# セル追記列位置
k = 10

url = "https://www.bloomberg.co.jp/quote/"
url_add = ":US"
sheet = edit_file.edit_excel(sheet, url, url_add, i, j, k)
wb.save("D:\OneDrive\sbi\investment.xlsx")


print("-------------------------割合出力------------------------------")

sheet = wb["時価割合"]

# 銘柄取得列
j = 1
# セル取得開始位置
i = 2
# セル追記列位置
k = 5

sheet = edit_file.edit_excel(sheet, url, url_add, i, j, k)
wb.save("D:\OneDrive\sbi\investment.xlsx")

