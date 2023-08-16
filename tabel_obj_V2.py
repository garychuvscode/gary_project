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
        ind_range0=0,
        normal_random=0,
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
        self.ind_row = self.ran_shape[0]
        self.ind_col = self.ran_shape[1]

        self.nor_Row = ind_range0.row
        self.nor_Col = ind_range0.column

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
        val_return = self.ind_range(
            (self.nor_Row + ind_row0, self.nor_Col + ind_col0)
        ).value
        return val_return
        pass

    def set_value(self, ind_row0=0, ind_col0=0, val_set0=0):
        """
        set value with index_Row and index_Col
        """
        self.ind_range(
            (self.nor_Row + ind_row0, self.nor_Col + ind_col0)
        ).value = val_set0

        pass

    def copy_to(self, dest_range0=0):
        """
        copy this table to the destination range
        """

        dest_range0.value = self.ind_range.value

        pass

    def compare_to(self, dest_range0=0, table_b0=0, space0=3, diff=1):
        """
        copy this table and the other table to the destination range for
        comparison \n
        table_a (this table); table_b (input table); comparison_table
        comparison = table_a - table_b
        diff set to 1 => also build he comparison table

        """

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
