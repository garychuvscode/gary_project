import xlwings as xw

# 连接到Excel应用程序
app = xw.App(visible=False)  # 可以设置 visible=True 以可见方式启动Excel

# 打开工作簿
wb = app.books.open("your_workbook.xlsx")

# 获取源表格的左上角单元格，例如 A1
start_cell = wb.sheets["Sheet1"].range("A1")

# 使用 expand 方法来动态识别表格范围
g_range = start_cell.expand("table")

# 在目标工作簿的目标位置粘贴
g_range.api.Copy()
wb.sheets["Sheet2"].range("A1").api.PasteSpecial()

# 关闭Excel应用程序
app.quit()
