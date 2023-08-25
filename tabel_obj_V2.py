"""
this object is used to load the table from excel to python
by using normalization, table index and dimension were given
and all the element can be access from the object

normal table => with x and y index
random table => only load the index and using the coordination of element
to define

"""

# whatch out the index of excel is y,x ; which is row, column

import xlwings as xw
import win32api as win_ap


class table_gen:
    def __init__(
        self,
        ind_range0=0,
        normal_random0=0,
    ):
        """
        setup the table object by assign range variable in
        normal_random => 0-normal; 1-random
        normal => with x, y index
        random => instrument_ctrl or others
        """

        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            ind_range0 = self.sh_comp.range((1, 1))

        # record the range for this table
        self.ind_range = ind_range0
        # knowing the dimenstion and starting index
        self.ran_shape = ind_range0.shape
        self.ran_sheet = ind_range0.sheet
        self.ran_books = ind_range0.sheet.book

        # dimension of the range
        self.ind_row = self.ran_shape[0]
        self.ind_col = self.ran_shape[1]

        # starting index (normalize)
        self.nor_Row = ind_range0.row
        self.nor_Col = ind_range0.column

        # this is the index cell used for shifting
        self.nor_Cell = self.ran_sheet.range((self.nor_Row, self.nor_Col))

        # space setting for comparison table copy
        self.space = 3

        # type of table
        """
        normal(0) => with row and column index
        random(1) => all data
        """
        self.normal_ind = normal_random0

        # initial of table message print in the terminal
        print(
            f"""table object generated \n

            """
        )

        pass

    def get_value(self, ind_row0=0, ind_col0=0):
        """
        get value with index_Row and index_Col
        """
        val_return = self.ran_sheet.range(
            (self.nor_Row + ind_row0, self.nor_Col + ind_col0)
        ).value
        return val_return
        pass

    def set_value(self, ind_row0=0, ind_col0=0, val_set0=0):
        """
        set value with index_Row and index_Col
        """
        self.ran_sheet.range(
            (self.nor_Row + ind_row0, self.nor_Col + ind_col0)
        ).value = val_set0

        pass

    def copy_to(self, dest_range0=0):
        """
        copy this table to the destination range
        only the value
        """

        dest_range0.value = self.ind_range.value

        pass

    def compare_to(self, dest_range0=0, table_b0=0, space0=3, diff=1):
        """
        copy this table and the other table to the destination range for
        comparison \n
        table_a (this table); table_b (input table); comparison_table
        comparison = table_a - table_b
        diff set to 1 => also build he comparison table \n
        dimension is based on the table call function

        """

        prog_fake = 0
        if prog_fake == 1:
            table_b0 = table_gen()
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            dest_range0 = self.sh_comp.range((1, 1))

            pass

        # first to copy the original table to destination range
        self.copy_to(dest_range0=dest_range0)

        # build up the next range index for second table
        ind_cell2 = (
            dest_range0.row,
            dest_range0.column + self.ind_col + self.space,
        )
        ran_table2 = dest_range0.sheet.range(ind_cell2)

        table_b0.copy_to(dest_range0=ran_table2)

        if diff == 1 and self.normal_ind == 0:
            # generate new range index for the comparison table
            ind_cell3 = (ind_cell2[0], ind_cell2[1] + self.ind_col + self.space)
            ran_table_3 = dest_range0.sheet.range(ind_cell3)
            # copy a table for diff
            # this copy is for index update
            self.copy_to(dest_range0=ran_table_3)

            x_row = 0
            while x_row < (self.ind_row - 1):
                x_col = 0
                while x_col < (self.ind_col - 1):
                    # index add 1 to shift to element place

                    # source 1 table
                    ind_r0 = dest_range0.row + 1 + x_row
                    ind_c0 = dest_range0.column + 1 + x_col

                    # source 2 table
                    ind_r1 = ind_cell2[0] + 1 + x_row
                    ind_c1 = ind_cell2[1] + 1 + x_col

                    # output table
                    ind_r2 = ind_cell3[0] + 1 + x_row
                    ind_c2 = ind_cell3[1] + 1 + x_col

                    # save the difference to the end
                    dest_range0.sheet.range(ind_r2, ind_c2).value = float(
                        dest_range0.sheet.range(ind_r0, ind_c0).value
                    ) - float(dest_range0.sheet.range(ind_r1, ind_c1).value)

                    x_col = x_col + 1
                    pass

                x_row = x_row + 1
                pass

            pass

        pass

    def compare_to0(self, dest_range0=0, table_b0=0, space0=3, diff=1):
        """
        shift version of compare
        230818 - note: original solution should be better,
        but offset of range can be known how to implement
        """

        prog_fake = 0
        if prog_fake == 1:
            table_b0 = table_gen()
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            dest_range0 = self.sh_comp.range((1, 1))

            pass

        # first to copy the original table to destination range
        self.copy_to(dest_range0=dest_range0)

        # build up the next range index for second table
        ind_cell1_ran = dest_range0.sheet.range(dest_range0.row, dest_range0.column)
        ind_cell2_ran = ind_cell1_ran.offset(0, self.ind_col + self.space)

        # for using cell variable in offset, this is range, not array
        # ran_table2 = dest_range0.sheet.range(ind_cell2)

        ind_cell2 = (dest_range0.row, dest_range0.column + self.ind_col + self.space)

        table_b0.copy_to(dest_range0=ind_cell2_ran)

        if diff == 1 and self.normal_ind == 0:
            # generate new range index for the comparison table

            ind_cell3 = (ind_cell2[0], ind_cell2[1] + self.ind_col + self.space)

            ind_cell3_ran = ind_cell2_ran.offset(0, self.ind_col + self.space)

            # for using cell variable in offset, this is range, not array
            # ran_table_3 = dest_range0.sheet.range(ind_cell3)

            # copy a table for diff
            # this copy is for index update
            self.copy_to(dest_range0=ind_cell3_ran)

            x_row = 0
            while x_row < (self.ind_row - 1):
                x_col = 0
                while x_col < (self.ind_col - 1):
                    # index add 1 to shift to element place

                    # source 1 table
                    ind_r0 = dest_range0.row + 1 + x_row
                    ind_c0 = dest_range0.column + 1 + x_col

                    # source 2 table
                    ind_r1 = ind_cell2[0] + 1 + x_row
                    ind_c1 = ind_cell2[1] + 1 + x_col

                    # output table
                    ind_r2 = ind_cell3[0] + 1 + x_row
                    ind_c2 = ind_cell3[1] + 1 + x_col

                    # save the difference to the end
                    dest_range0.sheet.range(ind_r2, ind_c2).value = float(
                        dest_range0.sheet.range(ind_r0, ind_c0).value
                    ) - float(dest_range0.sheet.range(ind_r1, ind_c1).value)

                    x_col = x_col + 1
                    pass

                x_row = x_row + 1
                pass

            pass

        pass

    def ran_shift(self, shift_row0=0, shift_col0=0):
        """
        this function help to shift the range

        please ask chat GPT for the action need to implement first
        so you will know if this function is needed to build by yourself or not

        chat GPT generation as below:

        Q: xlwings 如何調整 range 變數 single cell 的 row column 參數?

        A : 使用 xlwings, 您可以使用 range 物件的 .resize 方法來調整範圍的大小。
        要調整 range 物件的單個單元格的 row 和 column 參數，您可以使用 .offset 方法。
        下面是一個示例，說明如何使用 xlwings 來調整單個單元格的 row 和 column 參數：

        ======

        import xlwings as xw

        # 開啟 Excel 文件
        wb = xw.Book('your_file.xlsx')
        # 選取工作表
        sheet = wb.sheets['Sheet1']

        # 選取 B2 單元格
        cell = sheet.range('B2')

        # 調整單元格的 row 和 column 參數
        new_cell = cell.offset(row_offset=2, column_offset=3)

        # 輸出新單元格的地址
        print(new_cell.address)

        ======

        在這個例子中，我們選取了 B2 單元格，然後使用 .offset 方法將 row 和 column
        分別增加了 2 和 3,最後輸出新單元格的地址。輸出的結果將是 E4。

        """

        pass

    def get_range(self):

        # 獲取當前活動的Excel應用程序
        app = xw.apps.active

        # 獲取當前視窗的選定範圍
        selected_range = xw.Range(app.selection.address)

        # 使用selected_range進行後續操作
        # 例如，可以使用selected_range.value獲取範圍的值

        return selected_range

    pass


if __name__ == "__main__":
    #  the testing code for this file object

    book_name = "test"
    control_book_trace = "c:\\py_gary\\test_excel\\" + str(book_name) + ".xlsm"

    wb_test = xw.Book(control_book_trace)
    ran_ini = wb_test.sheets("Sheet1").range((2, 2), (6, 6))
    sh_test = wb_test.sheets("Sheet2")

    table_test = table_gen(ind_range0=ran_ini)

    test_index = 1

    if test_index == 0:
        # get value, set value, and copy table

        a = table_test.get_value(ind_row0=5, ind_col0=0)
        print(f"the value got is {a}")

        table_test.set_value(ind_row0=4, ind_col0=4, val_set0=3)

        ran_dest = sh_test.range((3, 3))

        table_test.copy_to(dest_range0=ran_dest)

        ran_dest = sh_test.range((3, 3))

        x = 0
        while x < 2:
            x = x + 1
            pass

        pass

    elif test_index == 1:
        # testing for comparison table copy
        ran_compared_table = wb_test.sheets("Sheet1").range((3, 3), (7, 7))
        table_to_be_compare = table_gen(ran_compared_table)
        ran_dest = wb_test.sheets("Sheet2").range((2, 2))

        table_test.compare_to(ran_dest, table_to_be_compare)

        pass

    elif test_index == 3 :
        # build up the loop for copying the table or generate the comparison table



        gen_operation = 1
        while gen_operation < 2 :

            # ask user to select the table for copy and generate the copied to a new range
            content_str = 'select table for the copy and press ok'
            title_str = 'select table 1'
            box_type = 0
            win_ap.MessageBox(0, content_str, title_str, box_type)


        pass

    pass
