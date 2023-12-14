import xlwings as xw


class ExcelProcessor:
    def __init__(self):
        self.x_axis = None
        self.y_axis = None

    def process_range(self, start_cell):
        # 使用 expand 方法展开范围，指定 direction='right' 表示水平方向展开
        expanded_range = start_cell.expand("right")

        # 获取 x 轴的参数
        self.x_axis = expanded_range.rows[0].value

        # 使用 expand 方法展开范围，指定 direction='down' 表示垂直方向展开
        expanded_range = start_cell.expand("down")

        # 获取 y 轴的参数
        self.y_axis = expanded_range.columns[0].value


# 连接到Excel应用程序
app = xw.App(visible=True)

# 打开工作簿
wb = app.books.open("your_workbook.xlsx")

# 获取要展开的起始单元格
start_cell = wb.sheets["Sheet1"].range("A1")

# 创建 ExcelProcessor 类的实例
excel_processor = ExcelProcessor()

# 调用 process_range 方法
excel_processor.process_range(start_cell)

# 输出 x 轴和 y 轴的参数
print("X Axis:", excel_processor.x_axis)
print("Y Axis:", excel_processor.y_axis)

# 关闭Excel应用程序
app.quit()
