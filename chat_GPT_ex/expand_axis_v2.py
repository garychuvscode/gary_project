import xlwings as xw

# 连接到Excel应用程序
app = xw.App(visible=False)  # 可以设置 visible=True 以可见方式启动Excel

# 打开工作簿
wb = app.books.open("your_workbook.xlsx")

# 获取源表格的左上角单元格，例如 A1
start_cell = wb.sheets["Sheet1"].range("A1")

# 使用 expand 方法来动态识别表格范围
g_table = start_cell.expand("table")

# 获取表格范围的值
table_values = g_table.value

# 获取 x 轴和 y 轴
x_axis = table_values[0]  # 假设第一行是 x 轴
y_axis = [row[0] for row in table_values]  # 假设第一列是 y 轴

# 在目标工作簿的目标位置粘贴
g_table.api.Copy()
wb.sheets["Sheet2"].range("A1").api.PasteSpecial()

# 关闭Excel应用程序
app.quit()

# 现在 x_axis 和 y_axis 可以用于进一步处理
print("X Axis:", x_axis)
print("Y Axis:", y_axis)
