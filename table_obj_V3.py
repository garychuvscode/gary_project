# turn off the formatter
# fmt: off

'''
this object is used to load table into object
by different way to define the table and provide different way to process
1. get or set value
2. table information (where is this table from)
(contain the shape, sheet, book, and file trace)
3. table comparison or plot
4. column or row access for different comparison method
(need to use column obj? or not? => need to think about)

'''

# whatch out the index of excel is y,x ; which is row, column

import xlwings as xw
import win32api as win_ap

# pandas and matplot is data science's good friend~
# import pandas as pd


class table_gen():
    def __init__(self, **kwargs):
        '''
        table object:
        for table index cell input: 'ind_row', 'ind_column'
        => this is single cell and auto scan for table

        '''

        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            ind_range0 = self.sh_comp.range((1, 1))
            self.ind_range = ind_range0

        # default table type is 0, using index cell input and get table by expand
        self.table_type = 0

        # 定义要匹配的关键字
        # watchout this dict can only have keys, don't causing error
        self.ctrl_para_dict = {'ind_cell', 'all_range', 'x_length', 'y_length', 'table_type' }
        # 初始化变量
        self.ind_cell = self.all_range = self.x_length = self.y_length = None


        # 遍历 kwargs 中的键值对
        for key, value in kwargs.items():
            # 如果键在 valid_keys 中，将对应的值存入相应的变量中
            if key in self.ctrl_para_dict:
                # setattr(self, f"ind_{key}", value)
                setattr(self, key, value)

                # tabletype should update automatically if input
                pass
            pass

        # 打印结果
        print(f'table create finished')
        print(f"ind_cell: {self.ind_cell}")
        print(f"all_range: {self.all_range}")
        print(f"x_length: {self.x_length}")
        print(f"y_length: {self.y_length}")
        print(f"table_type: {self.table_type}")

        # 231204 other type of table will be added after there are request

        # record the range for this table
        self.ind_range = self.ind_cell.expand("table")
        # knowing the dimenstion and starting index
        self.t_shape = self.ind_range.shape
        self.t_sheet = self.ind_range.sheet
        self.t_book = self.t_sheet.book
        self.t_file_trace = self.t_book.fullname

        # 231006 include the app_name of what parameter load is using
        self.app_org = self.t_book.app
        self.app_name = self.t_book.app._pid
        print(self.app_org.books)


        # dimension of the range
        self.row_dim = self.t_shape[0]
        self.col_dim = self.t_shape[1]

        # starting index (normalize)
        self.nor_Row = self.ind_range.row
        self.nor_Col = self.ind_range.column


        table_values = self.ind_range.value

        # 获取 x 轴和 y 轴
        x_axis = table_values[0]  # 假设第一行是 x 轴
        y_axis = [row[0] for row in table_values]  # 假设第一列是 y 轴

        # 使用 expand 方法展开范围，指定 direction='right' 表示水平方向展开
        expanded_range = self.ind_range.expand("right")

        # 获取 x 轴的参数
        self.x_axis = expanded_range.rows[0].value

        # 使用 expand 方法展开范围，指定 direction='down' 表示垂直方向展开
        expanded_range = self.ind_range.expand("down")

        # 获取 y 轴的参数
        self.y_axis = expanded_range.columns[0].value

        pass

    def table_output(self, to_range0=0):
        '''
        output this object to a range (index cell)
        '''
        if to_range0 != 0 :

            pass

    def get_row(self):


        return 0

    def get_col(self):


        return 0


if __name__ == "__main__":
    #  the testing code for this file object

    book_name = "test"
    control_book_trace = "c:\\py_gary\\test_excel\\" + str(book_name) + ".xlsm"

    wb_test = xw.Book(control_book_trace)
    input_ind_cell = wb_test.sheets('Sheet3').range((3,3))

    table_test = table_gen(ind_cell=input_ind_cell)

    test_index = 1

    if test_index == 0:


        pass

    pass
