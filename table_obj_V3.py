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
        default assume there are x and y axis, number process will exclude axis range
        set table_type 1 for the case of pure vaule table
        '''

        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            ind_range0 = self.sh_comp.range((1, 1))
            self.ind_range = ind_range0

        # default table type is 0, using index cell input and get table by expand
        # table type = 1 is the input table don't contain axis
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
        # 231206 these information can used to let people know where does the table from
        # knowing the condition

        # record the range for this table
        self.ind_range = self.ind_cell.expand("table")
        # create the pure value index and pure value range to make process easier
        # 231206 added fro the copy and number process of the table result
        self.ind_cell_no_axis = self.ind_cell.offset(1,1)
        self.ind_range_no_axis = self.ind_cell_no_axis.expand("table")
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


        # table_values = self.ind_range.value
        # # 获取 x 轴和 y 轴
        # x_axis = table_values[0]  # 假设第一行是 x 轴
        # y_axis = [row[0] for row in table_values]  # 假设第一列是 y 轴

        # 使用 expand 方法展开范围，指定 direction='right' 表示水平方向展开
        # 231206: ind_cell.expand("right") can also be use here, and may no need for rows[0]
        expanded_range = self.ind_range.expand("right")
        # 获取 x 轴的参数 (this it a list, not a range)
        self.x_axis = expanded_range.rows[0].value
        self.x_axis_range = expanded_range.rows[0]
        self.x_axis_len = len(self.x_axis)

        # 使用 expand 方法展开范围，指定 direction='down' 表示垂直方向展开
        expanded_range = self.ind_range.expand("down")

        # 获取 y 轴的参数 (this it a list, not a range)
        self.y_axis_range = expanded_range.columns[0]
        self.y_axis = expanded_range.columns[0].value
        # 231206: it's easy to get the value from column, but not easy to put it back
        self.y_axis_len = len(self.y_axis)

        # 输出原始表格范围的地址
        print(f"Original range: {self.ind_range.address}")


        pass

    def table_output(self, to_cell0=0, only_value0=1, format0='0.00%'):
        '''
        output this object to a range (index cell)
        should contain table information at left of table
        input the index_cell contain x and y axis
        format0: 2-dig => 0.00 ; 2-dig% => 0.00% ; 4-dig => 0.0000
        '''
        if to_cell0 != 0 :
            if only_value0 == 0 :
                # all property
                self.ind_range.copy(destination=to_cell0, expand='table')
            else:
                # build up the axis to cell index reference

                # 231207 try a new method
                # tmp_range = to_cell0.expand('down').resize(self.y_axis_len,1)
                # tmp_range.value = self.y_axis_range.value
                self.paste_col(src_col=self.y_axis_range, dest_col=to_cell0)

                # x-axis can opy directly
                tmp_range = to_cell0.expand('right').resize(1,self.x_axis_len)
                tmp_range.value = self.x_axis_range.value
                # 遍历源范围中的每个单元格，尝试将其转换为数字，否则将其设置为0
                to_cell0_no_axis = to_cell0.offset(1,1)
                # below line have issue of expand => no element beside for the destination range, change the expand to resize
                for src_cell, dest_cell in zip(self.ind_range_no_axis, to_cell0_no_axis.resize(self.y_axis_len-1, self.x_axis_len-1)):
                # for src_cell, dest_cell in zip(self.ind_range_no_axis, to_cell0_no_axis.expand('table')):
                    try:
                        # 检查源单元格的值是否为数字
                        if isinstance(src_cell.value, (int, float)):
                            # 设置小数位数为2位
                            # src_cell.number_format = format0
                            # 复制源单元格的值到目标单元格
                            dest_cell.value = src_cell.value
                            dest_cell.number_format = format0
                            pass
                        else:
                            # 尝试将源单元格的值转换为数字
                            value_as_number = float(src_cell.value)
                            # 将转换后的值复制到目标单元格
                            dest_cell.value = value_as_number
                            dest_cell.number_format = format0
                            pass

                        pass

                    except (ValueError, TypeError) as e:
                        # 打印错误类型和具体错误信息
                        dest_cell.value = 0
                        print(f'ERROR ELEMENT: {src_cell}')
                        print(f'transfer_err: {e}')
                        pass

                    pass

            pass

    # 231206: get_col and get_row merge to find_col_row

    def paste_col(self, src_col=0, dest_col=0):
        '''
        since there are unknow error for the paste column range at the xlwings
        process by this function using stupid method
        '''
        # cover by try, except to prevent crash by wrong input of function
        try:
            c_element = len(src_col)
            # auto adjustment the input dest range to match the source column size
            dest_col = dest_col.resize(c_element,1)
            x_element = 0
            while x_element < c_element :
                dest_col.rows[x_element].value = src_col.rows[x_element].value
                print(f'tansfer from {src_col} to {dest_col}, with element {x_element} and content {src_col.rows[x_element].value}')
                x_element = x_element + 1
                pass

            # return the dest_col range variable

            '''
            # also can try this code if needed (from chat GPT)
            for x_element, src_value in enumerate(src_col.rows):
                dest_col.rows[x_element].value = src_value
                print(f'transfer from {src_col} to {dest_col}, with element {x_element} and content {src_value}')
            '''

            return dest_col
            pass

        except Exception as e:
            print(f'there are error {e} cause by the paste column during operation')

            pass

        pass

    def index_shift(self, sh_row0=0, sh_col0=0):
        '''
        shift the index cell and also the dimension of table
        return table object
        '''
        new_cell = self.ind_cell.offset(sh_row0, sh_col0)

        table_new = table_gen(ind_cell=new_cell)

        return table_new

    def find_col_row(self, find_content0='', row_col0='col', obj_name0='grace', extra_fun0=0):
        '''
        input the axis index to look for, default set to find col
        extra_fun0=0 (default) => george's PMU eff index correction
        extra_fun0=1 => gary's eff index correction
        '''
        if row_col0 == 'col':
            # search for the x-axis to get column
            res_ind = self.search_element(source_list0=self.x_axis, search_content0=find_content0)
            ind_cell_new = self.ind_cell.offset(0, res_ind)
            temp_data = data_obj(ind_cell=ind_cell_new, row_col=row_col0, axis='y')
            # put the data and axis index to the data self.x_axis.valueobject
            temp_data.axis_content = self.y_axis
            temp_data.data_content = self.ind_range.columns[res_ind].value
            temp_data.data_len = len(temp_data.data_content)

            # value index correction
            if extra_fun0 == 0:
                # GPL_PMU efficiency
                t_cell = self.ind_cell.offset(1, -1)
                temp_data.axis_content[0] = t_cell.value
            elif extra_fun0 == 1:
                # gary_eff
                t_cell = self.ind_cell.offset(0, 0)
                temp_data.axis_content[0] = t_cell.value


        if row_col0 == 'row':
            # search for the x-axis to get row
            res_ind = self.search_element(source_list0=self.y_axis, search_content0=find_content0)
            ind_cell_new = self.ind_cell.offset(res_ind, 0)
            temp_data = data_obj(ind_cell=ind_cell_new, row_col=row_col0, axis='y')
            # put the data and axis index to the data self.x_axis.valueobject
            temp_data.axis_content = self.x_axis
            temp_data.data_content = self.ind_range.rows[res_ind].value
            temp_data.data_len = len(temp_data.data_content)

            if extra_fun0 == 0:
                # GPL_PMU efficiency
                t_cell = self.ind_cell.offset(-1, 1)
                temp_data.axis_content[0] = t_cell.value
            elif extra_fun0 == 1:
                # gary_eff
                t_cell = self.ind_cell.offset(-1, 1)
                temp_data.axis_content[0] = t_cell.value

        print(f'finished finding the row or col, item {find_content0}, with index {res_ind}')
        print(f'the axis: {temp_data.axis_content}, and data: {temp_data.data_content}')
        temp_data.data_obj_name = str(obj_name0)

        return temp_data

    def search_element(self, source_list0=0, re_type0='ind', search_content0=''):
        '''
        ** this function only return the first item, need to change code for muti return
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

        pass

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
        self.data_len = 0

        pass

    def status_chk(self):
        '''
        used to check object status, output the content in terminal
        '''

        print(f'''now the name is {self.data_obj_name},
index cell {self.ind_cell}, len = {self.data_len}
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
            print(f'the name assigned to default: {self.data_obj_name}')
            pass
        else:
            self.data_obj_name = str(name0)
            print(f'the name assigned by external : {self.data_obj_name}')
        pass


class chart_obj():

    def __init__(self, ind_cell0=0, ind_cell_d0=0, dimx0=0, dimy0=0, c_name0='Grace', x_name0='weekday', y_name0='cute%'):
        '''
        ind_cell0: index cell, in type 'range' for chart
        ind_cell_d0: index cell for data table from cell to range
        dimx0, dimy0: the dimensiont of chart, can use default (row_count, column count)
        axis: contain axis info or not, 'y', 'n'
        '''

        # import the initialization parameter from the definition
        self.ind_cell = ind_cell0
        self.ind_sheet = ind_cell0.sheet

        # 231209 can use the range.top; rnage.left to get settings
        # # chart location: left=row, top=col
        # row_count = self.ind_cell.row
        # col_count =self.ind_cell.column
        # # this is the fixed constant, won't change
        # row_u = 48
        # col_u = 17.5
        # # finall setting for left and top
        # self.c_row = row_u * row_count-1
        # self.c_col = col_u * col_count-1


        self.ind_cell_d = ind_cell_d0
        self.ind_cell_d_n_axis = self.ind_cell_d.offset(1,1)
        self.ind_cell_d_col = self.ind_cell_d.offset(0,1)
        # for width and heigh, use default if no input
        if dimy0 == 0:
            self.dimx = 355
        else:
            self.dimx = dimx0
        if dimy0 == 0:
            self.dimy = 211
        else:
            self.dimy = dimy0

        # x is all means x-axis here, inthe data table is down,
        # so use down for x
        self.x_len = len(self.ind_cell_d_n_axis.expand('down'))
        self.y_len = len(self.ind_cell_d_n_axis.expand('right'))

        self.c_name = c_name0
        self.x_name = x_name0
        self.y_name = y_name0

        # self.x_range = ind_cell0
        # self.y_range = ind_cell0

        tmp_cell = self.ind_cell_d.offset(1,0)
        self.x_axis = tmp_cell.expand("down")

        # this chart type is set to default
        self.chart_type = "xy_scatter_smooth_no_markers"

        # refer more chart setting for
        # https://docs.xlwings.org/en/stable/api/chart.html

        # setting of the en/dis of the minjor axis in chart
        # change this item befor chart add, and this can be enable
        self.en_minjor_axis = False
        self.en_max_min_auto_change = False


        pass

    def g_chart_add(self):
        '''
        add the chart based on object information
        '''
        # insert the chart
        self.ch_obj = self.ind_sheet.charts.add(left=self.ind_cell.left,top=self.ind_cell.top,width=self.dimx,height=self.dimy)
        self.ch_obj.set_source_data(self.ind_cell_d_col.expand('table'))
        self.ch_obj.chart_type = self.chart_type
        self.ch_obj.name = self.c_name
        self.ch_obj.api[1].HasTitle=True        #圖表有標題
        self.ch_obj.api[1].ChartTitle.Text=self.c_name  #標題文字
        self.ch_obj.api[1].Axes(1).HasTitle = True # This line creates the X axis label.
        self.ch_obj.api[1].Axes(2).HasTitle = True # This line creates the Y axis label.
        self.ch_obj.api[1].Axes(1).AxisTitle.Text = self.x_name
        self.ch_obj.api[1].Axes(2).AxisTitle.Text = self.y_name

        '''
        Axes(2) setting is in constant in xlwings:

        class AxisType:
        xlCategory = 1  # from enum XlAxisType
        xlSeriesAxis = 3  # from enum XlAxisType
        xlValue = 2  # from enum XlAxisType

        usually call by:

        Axes(xw.constants.AxisType.xlSeriesAxis) 是用來訪問
        Excel 圖表的系列軸(Series Axis)的方法。系列軸通常用於
        顯示圖表中的不同數據系列。在某些類型的圖表中，你可能會
        看到一條線或軸，該線或軸用來標記或切換數據系列。這與 Axes
        (xw.constants.AxisType.xlCategory)（類別軸，通常是 x 軸）
        和 Axes(xw.constants.AxisType.xlValue)（數值軸，通常是
        y 軸）不同。系列軸通常出現在堆疊的圖表或其他需要區分不同
        系列的情況下。

        C:\\py_gary\\py_virtual_3p10\\Lib\\site-packages\\xlwings

        '''

        # config the x-axis of each line
        for i in range (self.y_len) :
            self.ch_obj.api[1].SeriesCollection(i+1).XValues = self.x_axis.api
            print(f'change x-axis setting for curve {i}')

        # for the max and min of chart, default use 105% of max and 90% of min


        # after the auto generation of the chart, disable the auto adjustment

        # 231211: no need to disable, just just follow the sequence => major_uint, min, max should be ok
        # # == this decide the en/dis of the minor axis
        # self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).HasMinorGridlines = self.en_minjor_axis
        # self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).HasMinorGridlines = self.en_minjor_axis
        # == this disable auto change of maximum and minimum
        # self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).AutomaticMaximumScale = self.en_max_min_auto_change
        # self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).AutomaticMinimumScale = self.en_max_min_auto_change

        # 231211 testing code
        self.x_axis_major_u(x_major0=0.01)
        # self.x_axis_minor_u(x_minor0=1)
        self.y_axis_major_u(y_major0=0.02)
        # self.y_axis_minor_u(y_minor0=0.01)


        # control the amount of line within 10, otherwise there are display error
        self.x_axis_min(x_min0=0.00)
        self.x_axis_max(x_max0=0.05)
        self.y_axis_min(y_min0=0.85)
        self.y_axis_max(y_max0=0.95)



    def chart_display(self, x_y0='x', max0=0, min0=0, transfer=5):
        '''
        generate of chart will be auto axis, and it's able to change
        after the chart finished
        x_y0 => x, y selection
        the major unit will be define as (max0 - min0)/transfer
        '''
        major_u = ( max0 - min0 ) / transfer
        if x_y0 == 'x':
            # change x axis
            self.x_axis_major_u(x_major0=major_u)
            self.x_axis_min(x_min0=min0)
            self.x_axis_max(x_max0=max0)
        if x_y0 == 'y':
            # change y axis
            self.y_axis_major_u(y_major0=major_u)
            self.y_axis_min(y_min0=min0)
            self.y_axis_max(y_max0=max0)

        pass





    def x_axis_major_u(self, x_major0=0):
        '''
        change the x_axis major unit
        '''
        read_res = 0

        if x_major0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MajorUnit = x_major0
            print(f'x_maj_u set to {x_major0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MajorUnit
            print(f'x_maj_u is {read_res}')

        return read_res

    def x_axis_minor_u(self, x_minor0=0):
        '''
        change the x_axis minor unit
        the minor is default in disable
        '''
        read_res = 0

        if x_minor0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MinorUnit = x_minor0
            print(f'x_min_u set to {x_minor0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MinorUnit
            print(f'x_min_u is {read_res}')

        return read_res

    def y_axis_major_u(self, y_major0=0):
        '''
        change the y_axis major unit
        '''
        read_res = 0

        if y_major0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MajorUnit = y_major0
            print(f'y_maj_u set to {y_major0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MajorUnit
            print(f'y_maj_u is {read_res}')

        return read_res

    def y_axis_minor_u(self, y_minor0=0):
        '''
        change the y_axis minor unit
        the minor is default in disable
        '''
        read_res = 0

        if y_minor0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MinorUnit = y_minor0
            print(f'y_min_u set to {y_minor0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MinorUnit
            print(f'y_min_u is {read_res}')

        return read_res

    def x_axis_max(self, x_max0=0):
        '''
        change the x_axis maximum diplay
        '''
        read_res = 0

        if x_max0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MaximumScale  = x_max0
            print(f'x_max set to {x_max0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MaximumScale
            print(f'x_max is {read_res}')

        return read_res

    def x_axis_min(self, x_min0=0):
        '''
        change the x_axis minimum diplay
        '''
        read_res = 0

        if x_min0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MinimumScale  = x_min0
            print(f'x_min set to {x_min0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlCategory).MinimumScale
            print(f'x_min is {read_res}')

        return read_res

    def y_axis_max(self, y_max0=0):
        '''
        change the y_axis maximum diplay
        '''
        read_res = 0

        if y_max0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MaximumScale  = y_max0
            print(f'y_max set to {y_max0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MaximumScale
            print(f'y_max is {read_res}')

        return read_res

    def y_axis_min(self, y_min0=0):
        '''
        change the y_axis minimum diplay
        '''
        read_res = 0

        if y_min0 != 0 :
            self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MinimumScale  = y_min0
            print(f'y_min set to {y_min0}')

        else:
            read_res = self.ch_obj.api[1].Axes(xw.constants.AxisType.xlValue).MinimumScale
            print(f'y_min is {read_res}')

        return read_res







if __name__ == "__main__":


    #  the testing code for this file object

    book_name = "test"
    control_book_trace = "c:\\py_gary\\test_excel\\" + str(book_name) + ".xlsm"

    wb_test = xw.Book(control_book_trace)
    input_ind_cell = wb_test.sheets('Sheet3').range((3,3))

    table_test = table_gen(ind_cell=input_ind_cell)

    test_index = 5
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

        col_obj = table_test.find_col_row(find_content0='12', row_col0='col', obj_name0='col_1')
        print(f'the axis: {col_obj.axis_content} with data: {col_obj.data_content}')

        row_obj = table_test.find_col_row(find_content0='3.4', row_col0='row', obj_name0='row_1')
        print(f'the axis: {row_obj.axis_content} with data: {row_obj.data_content}')

        col_obj.name_ini()
        row_obj.status_chk()

        pass

    if test_index == 3:
        # testing for table copy

        new_ind = wb_test.sheets('Sheet3').range((34,15))

        table_test.table_output(to_cell0=new_ind, only_value0=1, format0='0.00%')
        pass

    if test_index == 4 :

        tmp_sh = wb_test.sheets('Sheet3')
        tmp_sh.charts.api
        a = tmp_sh.charts[0]
        b = tmp_sh.charts[1]
        print(tmp_sh.charts[0])
        print(tmp_sh.charts[1])
        row_count = 23
        col_count =10
        row_u = 48
        col_u = 17.5
        row = row_u * row_count-1
        col = col_u * col_count-1
        tmp_sh.charts.add(left=row, top=col)
        t_range = tmp_sh.range((30,20))
        print(f'top: {t_range.top}, left: {t_range.left}')
        # 设置 x 轴标题
        # a.api[1].Axes(2).HasTitle = True # This line creates the Y axis label.
        a.api[1].Axes(1).AxisTitle.Text = "x axis text"
        a.api[1].Axes(2).AxisTitle.Text = "y axis text"

    if test_index == 5 :

        tmp_sh = wb_test.sheets('Sheet3')
        ind_cell_in = tmp_sh.range((30, 10))
        ind_cell_d_in = tmp_sh.range((3, 3))
        ch_obj1 = chart_obj(ind_cell0=ind_cell_in, ind_cell_d0=ind_cell_d_in)
        ch_obj1.g_chart_add()

        ch_obj1.chart_display(x_y0='x', max0=3, min0=2.5)
        ch_obj1.chart_display(x_y0='y', max0=0.9, min0=0.7)


    # end of test mode
    pass
