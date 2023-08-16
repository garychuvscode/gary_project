import xlwings as xw

# 開啟 Excel 檔案
wb = xw.Book()

# 選取一個範圍
range_ = wb.sheets["工作表1"].range("A1:C3")

# 取得範圍的形狀
shape = range_.shape

a = range_.sheet
b = range_.sheet.book

# 取得範圍的行索引和列索引
rows = shape[0]  # 行數
columns = shape[1]  # 列數

print(f"範圍的行索引:1 至 {rows}")
print(f"範圍的列索引:1 至 {columns}")

# ===========

# 获取工作表
sheet = wb.sheets["工作表1"]

# 获取单元格的行列位置
# cell = sheet.range("A1:B2")

# for the range input not only 1 cell, the left up one is the index
cell = sheet.range("B3:C10")
row = cell.row
column = cell.column

print(f"单元格的行位置：{row}")
print(f"单元格的列位置：{column}")
