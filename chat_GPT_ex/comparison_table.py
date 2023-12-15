import xlwings as xw


class TableGen:
    def __init__(self, **kwargs):
        # 初始化变量...
        pass

    def table_comp(self, other_table):
        # 确保比较的表格是相同维度的
        if self.row_dim != other_table.row_dim or self.col_dim != other_table.col_dim:
            raise ValueError("表格维度不匹配")

        # 获取表格数据
        self_data = self.ind_range.value
        other_data = other_table.ind_range.value

        # 在这里添加您的比较逻辑，假设这里使用减法进行比较
        result_data = [
            [self_data[i][j] - other_data[i][j] for j in range(self.col_dim)]
            for i in range(self.row_dim)
        ]

        # 创建新的工作簿用于存放比较结果
        comp_wb = xw.Book()
        comp_sheet = comp_wb.sheets.add("Comparison_Result")

        # 将比较结果写入新的工作表
        comp_range = comp_sheet.range((1, 1)).expand("table")
        comp_range.value = result_data

        # 输出比较结果表格范围的地址
        print(f"Comparison result range: {comp_range.address}")

        # 返回新创建的工作簿和工作表
        return comp_wb, comp_sheet


# 示例用法
table1 = TableGen(ind_cell="A1", all_range="A1:C4")
table2 = TableGen(ind_cell="D1", all_range="D1:F4")

# 进行表格比较
comp_wb, comp_sheet = table1.table_comp(table2)

# 在这里可以继续处理 comp_wb 和 comp_sheet，例如保存工作簿
comp_wb.save("path_to_comparison_result.xlsx")
comp_wb.close()
