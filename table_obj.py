"""
this object is used to load the table from excel to python
by using normalization, table index and dimension were given
and all the element can be access from the object

normal table => with x and y index
random table => only load the index and using the coordination of element
to define

"""

import xlwings as xw


class table_gen:
    def __init__(
        self,
        trace0="",
        book_name0="",
        wb0=0,
        sheet_name0="",
        nor_x0=0,
        nor_y0=0,
        normal_random=0,
    ):
        """
        setup the table object by assign the trace or send
        the wb in object to define
        normal_random => 0-normal; 1-random
        normal => with x, y index
        random => instrument_ctrl or others
        nor_X0, nor_Y0 => table started index
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
        self.nor_X = nor_x0
        self.nor_Y = nor_y0

        pass

    def get_value(self, inx_x0=0, ind_y0=0):
        """
        get value with index_X and index_Y
        """
        val_return = self.main_sheet.range(
            (self.nor_X + inx_x0, self.nor_Y + ind_y0)
        ).value
        return val_return
        pass

    def set_value(self, inx_x0=0, ind_y0=0, val_set=0):
        """
        set value with index_X and index_Y
        """
        self.main_sheet.range(
            (self.nor_X + inx_x0, self.nor_Y + ind_y0)
        ).value = val_set

    pass


if __name__ == "__main__":
    #  the testing code for this file object

    pass
