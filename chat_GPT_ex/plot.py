import xlwings as xw
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import make_interp_spline


# fmt: off

class test_class():
    def __init__(self):
        pass


    # 定义 Excel 函数
    @xw.func
    @xw.arg("x", doc="X轴数据范围")
    @xw.arg("y", doc="Y轴数据范围")
    @xw.arg("x_name", doc="X轴名称")
    @xw.arg("y_name", doc="Y轴名称")
    @xw.ret(index=False, header=False)
    def create_smooth_scatter_plot(self, x=0, y=0, x_name="x_name", y_name="y_name"):
        # 将 Excel 范围转换为 NumPy 数组
        x_data = np.array(x.value)
        y_data = np.array(y.value)

        # 创建平滑曲线
        x_smooth = np.linspace(x_data.min(), x_data.max(), 300)
        spline = make_interp_spline(x_data, y_data, k=3)
        y_smooth = spline(x_smooth)

        # 创建散点图
        fig, ax = plt.subplots()
        ax.scatter(x_data, y_data, label="散点图", color="blue")
        ax.plot(x_smooth, y_smooth, label="平滑曲线", color="red")
        ax.set_xlabel(x_name)
        ax.set_ylabel(y_name)
        ax.legend()

        # 将图形嵌入到 Excel 中
        xw.Book('book1.xlsx').set_mock_caller()
        sheet = xw.Book.caller().sheets.active
        chart = sheet.charts.add()
        chart.set_source_data(sheet.range("A1:B{}".format(len(x_data) + 1)))
        chart.top = sheet.range("D16").top
        chart.left = sheet.range("D16").left
        chart.width = 400
        chart.height = 300

        # 显示图形
        plt.show()

        pass



if __name__ == "__main__":
    # testing of GPL MCU connection

    t_s = test_class()
    testing_index = 1

    # 连接到活动的 Excel 应用程序
    # app = xw.App()
    wb = xw.books('book1.xlsm')

    # 获取活动工作簿和活动工作表
    sheet = wb.sheets.active

    # 在活动工作表上命名 A1:A10 为 "X_Data"
    x_data_range = sheet.range("A1:A10")
    x_data_range.name = "X_Data"

    # 在活动工作表上命名 B1:B10 为 "Y_Data"
    y_data_range = sheet.range("B1:B10")
    y_data_range.name = "Y_Data"

    # 命名活动工作表为 "My_Sheet"
    sheet.name = "My_Sheet"

    # 关闭 Excel 应用程序
    # app.quit()

    if testing_index == 0:
        t_s.create_smooth_scatter_plot(
            x_data_range, y_data_range, x_data_range.name, y_data_range.name
        )

        pass

    elif testing_index == 1:

        def create_smooth_scatter_plot(x_data, y_data, x_name, y_name):
            # 将数据转换为 NumPy 数组
            x_data = np.array(x_data)
            y_data = np.array(y_data)

            # 创建平滑曲线
            x_smooth = np.linspace(x_data.min(), x_data.max(), 300)
            spline = make_interp_spline(x_data, y_data, k=3)
            y_smooth = spline(x_smooth)

            # 创建散点图
            fig, ax = plt.subplots()
            ax.scatter(x_data, y_data, label="散点图", color="blue")
            ax.plot(x_smooth, y_smooth, label="平滑曲线", color="red")
            ax.set_xlabel(x_name)
            ax.set_ylabel(y_name)
            ax.legend()

            # 显示图形
            plt.show()

        # 示例用法
        x_data = x_data_range.value
        y_data = y_data_range.value
        x_name = "X-axis"
        y_name = "Y-axis"

        create_smooth_scatter_plot(x_data, y_data, x_name, y_name)

        pass
