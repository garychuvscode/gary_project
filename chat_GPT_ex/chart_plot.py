import xlwings as xw


def create_chart(
    sheet_name,
    ind_cell,
    x_axis,
    y_axis,
    plot_type,
    x_window,
    y_window,
    x_name,
    y_name,
    x_dim,
    y_dim,
    plot_name,
    dest_cell,
):
    # 连接到Excel应用程序
    app = xw.App(visible=True)

    # 打开工作簿
    wb = app.books.open("your_workbook.xlsx")

    # 获取要作图的数据范围
    ws = wb.sheets[sheet_name]
    data_range = ws.range(ind_cell).expand("table")

    # 在工作表中创建图表
    chart = ws.shapes.add_chart(
        chart_type=plot_type, left=0, top=0, width=375, height=225, name=plot_name
    )

    # 设置图表数据范围
    chart.chart.set_source_data(data_range)

    # 设置图表标题
    chart.chart.chart_title.text = plot_name

    # 设置 x 轴标题
    x_axis_title_obj = chart.chart.axes(xw.constants.AxisCategory).axis_title
    x_axis_title_obj.text = x_name

    # 设置 y 轴标题
    y_axis_title_obj = chart.chart.axes(xw.constants.AxisValue).axis_title
    y_axis_title_obj.text = y_name

    # 设置 x 和 y 轴的窗口范围
    chart.chart.axes(xw.constants.AxisCategory).minimum_scale = x_axis - x_window
    chart.chart.axes(xw.constants.AxisCategory).maximum_scale = x_axis + x_window
    chart.chart.axes(xw.constants.AxisValue).minimum_scale = y_axis - y_window
    chart.chart.axes(xw.constants.AxisValue).maximum_scale = y_axis + y_window

    # 设置图表大小
    chart.width = x_dim
    chart.height = y_dim

    # 设置图表位置
    chart.left = dest_cell.left
    chart.top = dest_cell.top

    # 保存工作簿
    wb.save()

    # 关闭Excel应用程序
    app.quit()


# 调用函数
create_chart(
    sheet_name="Sheet1",
    ind_cell="A1",
    x_axis=2,
    y_axis=3,
    plot_type="xy_scatter_smooth_no_markers",
    x_window=1,
    y_window=1,
    x_name="X Axis",
    y_name="Y Axis",
    x_dim=400,
    y_dim=300,
    plot_name="MyChart",
    dest_cell=ws.range("V7"),
)
