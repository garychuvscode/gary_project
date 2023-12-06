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

        # # 获取 x 轴和 y 轴
        # x_axis = table_values[0]  # 假设第一行是 x 轴
        # y_axis = [row[0] for row in table_values]  # 假设第一列是 y 轴

        # 使用 expand 方法展开范围，指定 direction='right' 表示水平方向展开
        # 231206: ind_cell.expand("right") can also be use here, and may no need for rows[0]
        expanded_range = self.ind_range.expand("right")
        # 获取 x 轴的参数
        self.x_axis = expanded_range.rows[0].value

        # 使用 expand 方法展开范围，指定 direction='down' 表示垂直方向展开
        expanded_range = self.ind_range.expand("down")

        # 获取 y 轴的参数
        self.y_axis = expanded_range.columns[0].value
        # 231206: it's easy to get the value from column, but not easy to put it back 


        # 输出原始表格范围的地址
        print(f"Original range: {self.ind_range.address}")

        # 获取去除 x 轴后的新范围
        new_range_without_x = self.ind_range.resize(self.ind_range.rows.count, self.ind_range.columns.count - 1)

        # 输出新范围的地址
        print(f"New range without x-axis: {new_range_without_x.address}")

        # 获取去除 y 轴后的新范围
        new_range_without_y = self.ind_range.resize(self.ind_range.rows.count - 1, self.ind_range.columns.count)

        # 输出新范围的地址
        print(f"New range without y-axis: {new_range_without_y.address}")




        pass

    def table_output(self, to_cell0=0):
        '''
        output this object to a range (index cell)
        '''
        if to_cell0 != 0 :

            pass

    def get_row(self, row_ind=0):
        '''
        return the row range based on this table object
        use to mix the row data from different table
        row 0 may be the x axis
        '''
        # think about to add the y-axis to prevent issue of merge data

        return 0

    def get_col(self, ind_cell0=0, len0=0, axis0=0):
        '''
        return the column range based on this table object
        use to mix the column data from different table
        row 0 may be the x axis
        len0=0 => expand to the end
        axis0=0 => don't contain y-axis information
        '''
        # think about to add the y-axis to prevent issue of merge data

        return 0

    def index_shift(self, sh_row0=0, sh_col0=0):
        '''
        shift the index cell and also the dimension of table
        return table object
        '''
        new_cell = self.ind_cell.offset(sh_row0, sh_col0)

        table_new = table_gen(ind_cell=new_cell)

        return table_new

    def find_col_row(self, find_content0='', row_col0='col',obj_name0='grace'):
        '''
        input the axis index to look for, default set to find col
        '''
        if row_col0 == 'col':
            # search for the x-axis to get column
            res_ind = self.search_element(source_list0=self.x_axis, search_content0=find_content0)
            ind_cell_new = self.ind_cell.offset(0, res_ind)
            temp_data = data_obj(ind_cell=ind_cell_new, row_col=row_col0, axis='y')
            # put the data and axis index to the data self.x_axis.valueobject
            temp_data.axis_content = self.x_axis
            temp_data.data_content = self.ind_range.columns[res_ind].value

        if row_col0 == 'row':
            # search for the x-axis to get column
            res_ind = self.search_element(source_list0=self.y_axis, search_content0=find_content0)
            ind_cell_new = self.ind_cell.offset(res_ind, 0)
            temp_data = data_obj(ind_cell=ind_cell_new, row_col=row_col0, axis='y')
            # put the data and axis index to the data self.x_axis.valueobject
            temp_data.axis_content = self.y_axis
            temp_data.data_content = self.ind_range.rows[res_ind].value

        print(f'finished finding the row or col, item {find_content0}, with index {res_ind}')
        print(f'the axis: {temp_data.axis_content}, and data: {temp_data.data_content}')
        temp_data.data_obj_name = str(obj_name0)

        return temp_data

    def search_element(self, source_list0=0, re_type0='ind', search_content0=''):
        '''
        ** this function only return the dirst item, need to change code for muti return
        source_list0 => the list needto search
        re_type0 => 'ind' is return index ; 'item' is return content
        search_content0 => the string need to search
        '''

        # transfer to str list before comparison 
        # 使用列表推导式将列表中的所有元素转换为字符串
        my_string_list = [str(element) for element in source_list0]

        # 输出结果
        print(f"Original List: {source_list0}")
        print(f"String List: {my_string_list}")


        # 搜索的子字符串
        search_substring = str(search_content0)

        # 用于存储包含子字符串 '12' 的元素的列表
        result_elements = []

        # 用于存储包含子字符串 '12' 的元素的索引的列表
        result_indices = []

        # 遍历列表，记录包含子字符串 '12' 的元素和它们的索引
        for index, element in enumerate(my_string_list):
            '''
            在这个例子中,enumerate 函数用于同时获取元素和它们的索引。然后，
            使用列表推导式检查每个元素是否包含子字符串
            '12'，并返回包含该子字符串的元素的索引列表。
            '''
            try: 
                if search_substring in element:
                    result_elements.append(element)
                    result_indices.append(index)
                    pass 
                pass 
            except Exception as e:
                # show the exception but not stop the program
                # e is the exception message from program
                print(f"there are error {str(e)}, with element: {element}")

        # 输出结果
        print(f"The elements containing '{search_substring}' are: {result_elements}")
        print(f"The indices of elements containing '{search_substring}' are: {result_indices}")

        if re_type0 == 'ind':
            # return the index (only return the first one here)
            return int(result_indices[0])
        if re_type0 == 'item':
            # return the item contain searching content (also the first one)
            return result_elements[0]

class data_obj():

    def __init__(self, **kwargs):
        '''
        ind_cell: index cell, in type 'range'
        row_col: row or column data, 'row' or 'col'
        axis: contain axis info or not, 'y', 'n'
        '''
        # 定义要匹配的关键字
        # watchout this dict can only have keys, don't causing error
        self.ctrl_para_dict = {'ind_cell', 'row_col', 'axis'}
        # 初始化变量
        self.ind_cell = self.row_col = self.axis = None

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
        print(f"row_col: {self.row_col}")
        print(f"axis: {self.axis}")

        #== after random initializtion parameter:

        self.data_obj_name ='Grace'
        # this is the list for the data saving in data_obj
        self.data_content = []
        # this is the list for the axis index saving in data_obj
        self.axis_content = []

        pass

    def status_chk(self):
        '''
        used to check object status, output the content in terminal
        '''

        print(f'''now the name is {self.data_obj_name}, cell {self.ind_cell},
type: {self.row_col}, axis index: {self.axis} ;
the content: {self.data_content}
the axis index: {self.axis_content}
              ''')

        pass

    def name_ini(self, name0=0):
        '''
        update the object name after data input, use item 0 in content
        if there are no name input
        '''
        if name0 == 0 :
            self.data_obj_name = str(self.data_content[0])
            print(f'the name assigned: {self.data_obj_name}')
            pass
        pass




if __name__ == "__main__":
    #  the testing code for this file object

    book_name = "test"
    control_book_trace = "c:\\py_gary\\test_excel\\" + str(book_name) + ".xlsm"

    wb_test = xw.Book(control_book_trace)
    input_ind_cell = wb_test.sheets('Sheet3').range((3,3))

    table_test = table_gen(ind_cell=input_ind_cell)

    test_index = 2

    if test_index == 0:
        # teting for column copy and retrn column object


        # 获取要粘贴数据的起始单元格
        paste_start_cell = wb_test.sheets['Sheet4'].range('D1')
        paste_start_cell_2 = wb_test.sheets['Sheet4'].range('F1')


        # 假设有一行数据要粘贴
        data_to_paste_row = [1, 2, 3, 4, 5]

        # 将数据写入 Excel 列，使范围在垂直方向（下方）扩展
        paste_start_cell.expand('right').value = data_to_paste_row

        # 假设有一列数据要粘贴
        data_to_paste_column = [2, 2, 3, 4, 5]

        range_to_select_down = paste_start_cell.resize(5, 1)
        range_to_select_down.value = data_to_paste_column

        paste_start_cell_2.offset(len(data_to_paste_column), 0).value = data_to_paste_column

        # 假设有一个二维列表数据
        table_values = [[1, 'A', 10],
                        [2, 'B', 20],
                        [3, 'C', 30]]

        # 提取第一列的所有值
        first_column_values = [row[0] for row in table_values]

        # 将提取的值写入 Excel 一列（使用 expand 方法确保列的高度足够）
        paste_start_cell_2.expand('down').value = first_column_values

        # column_range = wb_test.sheets['Sheet3'].columns(3)

        tb2 = table_test.index_shift(sh_row0=2, sh_col0=6)


        pass

    if test_index == 1:
        # testing for getting x, y axis from the index cell

        y_axis = table_test.ind_cell.expand("down")
        y_value = y_axis.value

        x_axis = table_test.ind_cell.expand("right")
        x_value = x_axis.value

        pass

    if test_index == 2:
        # check 

        col_obj = table_test.find_col_row(find_content0='12', row_col0='col')
        print(f'the axis: {col_obj.axis_content} with data: {col_obj.data_content}')

        row_obj = table_test.find_col_row(find_content0='3.4', row_col0='row')
        print(f'the axis: {row_obj.axis_content} with data: {row_obj.data_content}')

        pass 


    # end of test mode 
    pass 
