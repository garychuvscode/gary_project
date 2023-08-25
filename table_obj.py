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


class table_gen:
    def __init__(
        self,
        trace0="",
        book_name0="",
        wb0=0,
        sheet_name0="",
        nor_row0=0,
        nor_col0=0,
        normal_random=0,
    ):
        """
        setup the table object by assign the trace or send
        the wb in object to define => need to be wb obj, not file name
        normal_random => 0-normal; 1-random
        normal => with x, y index
        random => instrument_ctrl or others
        nor_Row0, nor_Col0 => table started index
        """

        # loaded the work book to object
        if wb0 == 0:
            # using the trace
            self.control_book_trace = str(trace0) + str(book_name0) + ".xlsm"
            self.wb = xw.Book(self.control_book_trace)
            print(f"trace set to {self.control_book_trace}")

            pass

        else:
            self.wb = wb0
            pass

        self.main_sheet = self.wb.sheets(str(sheet_name0))
        self.nor_Row = nor_row0
        self.nor_Col = nor_col0

        pass

    def get_value(self, ind_row0=0, ind_col0=0):
        """
        get value with index_Row and index_Col
        """
        val_return = self.main_sheet.range(
            (self.nor_Row + ind_row0, self.nor_Col + ind_col0)
        ).value
        return val_return
        pass

    def set_value(self, ind_row0=0, ind_col0=0, val_set=0):
        """
        set value with index_Row and index_Col
        """
        self.main_sheet.range(
            (self.nor_Row + ind_row0, self.nor_Col + ind_col0)
        ).value = val_set

        pass

    pass


if __name__ == "__main__":
    #  the testing code for this file object

    wb_test = xw.books("obj_main.xlsm")

    table_test = table_gen(
        wb0=wb_test,
        sheet_name0="inst_ctrl",
        nor_row0=12,
        nor_col0=3,
        normal_random=1,
    )
    a = table_test.get_value(ind_row0=5, ind_col0=0)
    print(f"the value got is {a}")

    table_test.set_value(ind_row0=5, ind_col0=0, val_set=3)

    pass
