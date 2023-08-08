"""
this object used to handle all the report arragement after raw dat input
handling the combing of raw data from different sheet or area

try to make good summary to make report easier for others to understand
"""

# import related tool needed for report arragement
import xlwings as xw
import logging as log

# mainly for only process the report

# fmt: off

class report_arragement:
    def __init__(self, file_name0=""):
        """
        this class used to process some regular copy, paste and plot of report \n
        file input need to be .xlsm
        """
        # input file name of report
        self.file_name = str(file_name0)

        self.wb = xw.books(self.file_name)
        print(f"select {self.file_name} as report file")

        # name of sheet for example operation
        self.Buck_eff_example = "Eff_comp_ex"
        self.LDO_reg_example = "regulation_comp_LDO_ex"
        self.raw_sheet = "C_raw_NAB"
        self.sheet_sor = ""
        self.sheet_des = ""

        # name of source sheet
        self.source_sheet = 0
        self.destination_sheet = 0

        # spece between each table (column + space)
        self.space = 4

        pass

    def Buck_eff_load_regulation(self, ind_row=0, ind_col=0):
        """
        operate the copy form the table reference of row and
        column index input

        """

        # eff SY-0-1.9A
        ind1 = (5,5)
        ind2 = (43,4)
        row_c = 39
        col_c = 7

        sh_sor = self.wb.sheets(self.Buck_eff_example).range(5,3)
        sh_dest = self.Buck_eff_example

        pass

    def LDO_load_regulation(self):
        pass

    def table_comparison(self,ind_1=(1, 1),ind_2=(3, 3),row_count=3,col_count=3,ind_dest0=(5, 5),sheet_sor0="",sheet_des0="",diff=1):
        """
        copy table and generate comparison table
        diff = 1 => add the difference comparison
        """
        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            sh_sor = self.wb.sheets("temp")
            sh_des = self.wb.sheets("temp")

        ind_1_end = (ind_1[0] + row_count, ind_1[1] + col_count)
        ind_2_end = (ind_2[0] + row_count, ind_2[1] + col_count)

        if sheet_sor0 != '':
            self.sheet_sor = str(sheet_sor0)
            sh_sor = self.wb.sheets(self.sheet_sor)
            pass

        # if no sor0 input, set to default, may be error
        sh_sor = self.wb.sheets(self.sheet_sor)

        if sheet_des0 != '':
            self.sheet_des = str(sheet_des0)
            sh_des = self.wb.sheets(self.sheet_des)
            pass
        else:
            # if no des input, use the same sheet
            sh_des = sh_sor

        # sor_range().copy(res_range)

        # shift destination index
        ind_dest1 = (ind_dest0[0], ind_dest0[1] + col_count + self.space)
        ind_dest2 = (ind_dest0[0], ind_dest1[1] + col_count + self.space)

        sor_range = sh_sor.range(ind_1, ind_1_end)
        res_range = sh_des.range(ind_dest0)

        # copy table 1
        sor_range.copy(res_range)

        sor_range = sh_sor.range(ind_2, ind_2_end)
        res_range = sh_des.range(ind_dest1)

        # cpoy table 2
        sor_range.copy(res_range)

        sor_range = sh_sor.range(ind_2, ind_2_end)
        res_range = sh_des.range(ind_dest2)

        if diff == 1:

            # cpoy table diff
            sor_range.copy(res_range)

            x_row = 0
            while x_row < row_count:
                x_col = 0
                while x_col < col_count:
                    # index add 1 to shift to element place

                    # source 1 table
                    ind_r0 = ind_dest0[0] + 1 + x_row
                    ind_c0 = ind_dest0[1] + 1 + x_col

                    # source 2 table
                    ind_r1 = ind_dest1[0] + 1 + x_row
                    ind_c1 = ind_dest1[1] + 1 + x_col

                    # output table
                    ind_r2 = ind_dest2[0] + 1 + x_row
                    ind_c2 = ind_dest2[1] + 1 + x_col

                    # save the difference to the end
                    sh_des.range(ind_r2, ind_c2).value = (
                        sh_des.range(ind_r0, ind_c0).value - sh_des.range(ind_r1, ind_c1).value
                    )

                    x_col = x_col + 1
                    pass

                x_row = x_row + 1
                pass

        pass

    def table_copy(self,ind_1=(1, 1),row_count=3,col_count=3,ind_dest0=(5, 5),sheet_sor0="",sheet_des0=""):
        '''
        just copy one table, not for comparison
        '''
        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            sh_sor = self.wb.sheets("temp")
            sh_des = self.wb.sheets("temp")

        ind_1_end = (ind_1[0] + row_count, ind_1[1] + col_count)

        if sheet_sor0 != '':
            self.sheet_sor = str(sheet_sor0)
            sh_sor = self.wb.sheets(self.sheet_sor)
            pass

        # if no sor0 input, set to default, may be error
        sh_sor = self.wb.sheets(self.sheet_sor)

        if sheet_des0 != '':
            self.sheet_des = str(sheet_des0)
            sh_des = self.wb.sheets(self.sheet_des)
            pass
        else:
            # if no des input, use the same sheet
            sh_des = sh_sor

        # sor_range().copy(res_range)

        sor_range = sh_sor.range(ind_1, ind_1_end)
        res_range = sh_des.range(ind_dest0)

        # copy table 1
        sor_range.copy(res_range)

        pass

    def table_title(self):
        '''
        use this program to copy the title of table
        '''

        pass

    def sor_des_sheet_sel(self):
        '''
        select the sheet of sor and des
        '''
        pass


    pass




if __name__ == "__main__":
    #  the testing code for this file object

    test_index = 2
    rep = report_arragement(file_name0="test.xlsx")

    operation_index = 0
    file_name = "test.xlsx"
    rep2 = report_arragement(file_name0=file_name)

    if test_index == 0 :

        rep.table_comparison(
            ind_1=(5, 2),
            ind_2=(5, 9),
            row_count=7,
            col_count=4,
            ind_dest0=(15, 2),
            sheet_sor0="test_sh",
            sheet_des0='test_sh2',
            diff=1
        )

        pass

    elif test_index == 1 :
        rep.table_comparison(
            ind_1=(5, 2),
            ind_2=(5, 9),
            row_count=7,
            col_count=4,
            ind_dest0=(15, 2),
            sheet_sor0="test_sh",
            sheet_des0='',
            diff=1
        )
        pass

    elif test_index == 2 :

        rep.table_comparison(
            ind_1=(5, 2),
            ind_2=(5, 9),
            row_count=7,
            col_count=4,
            ind_dest0=(15, 2),
            sheet_sor0="test_sh",
            sheet_des0='',
            diff=0
        )

        pass



    pass
