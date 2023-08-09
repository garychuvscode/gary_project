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
    def __init__(self, file_name0="", full_trace0=0):
        """
        this class used to process some regular copy, paste and plot of report \n
        file input need to be .xlsm
        """
        self.full_trace = ''

        if full_trace0 == 0 :

            # input file name of report
            self.file_name = str(file_name0)

            self.wb = xw.books(self.file_name)
            print(f"select {self.file_name} as report file")
            self.full_trace = ''
        else:
            # file saving is based on the trace content
            self.wb = xw.Book(file_name0)
            self.full_trace = file_name0


        # name of sheet for example operation
        self.Buck_eff_example = "Eff_comp_ex"
        self.LDO_reg_example = "regulation_comp_LDO_ex"
        self.raw_sheet = "C_raw_NAB"

        # this is not sheet name, already sheet object
        self.sheet_sor = ""
        self.sheet_des = ""

        # name of source sheet
        self.source_sheet = 0
        self.destination_sheet = 0

        # spece between each table (column + space)
        self.space = 4

        pass

    def Buck_eff_load_regulation(self):
        """
        operate the copy form the table reference of row and
        column index input

        """

        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            self.sh_sy_eff_0p19 = self.wb.sheets("temp")

        self.sh_comp = self.wb.sheets(self.Buck_eff_example)

        self.buck_sh_name = self.sh_comp.range('C4').value
        self.sy_eff_0p19 = self.sh_comp.range('C5').value
        self.sy_eff_2 = self.sh_comp.range('C6').value
        self.nt_eff_0p19 = self.sh_comp.range('C8').value
        self.nt_eff_2 = self.sh_comp.range('C9').value

        # this is the index sheet for copy summary sheet to front of first sheet
        self.sh_sy_eff_0p19 = self.wb.sheets(self.sy_eff_0p19)


        sh_temp = self.sy_eff_0p19
        # eff SY-0-1.9A
        ind1 = (4,3)
        des = (43,4)
        row_c = 39
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)


        sh_temp = self.sy_eff_2
        # eff SY-2-8A
        ind1 = (4,3)
        des = (82,4)
        row_c = 29
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)


        sh_temp = self.nt_eff_0p19
        # eff NT-0-1.9A
        ind1 = (4,3)
        des = (43,16)
        row_c = 39
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)


        sh_temp = self.nt_eff_2
        # eff NT-2-8A
        ind1 = (4,3)
        des = (82,16)
        row_c = 29
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)


        sh_temp = self.sy_eff_0p19
        # load_reg SY-0-1.9A
        ind1 = (184,3)
        des = (142,4)
        row_c = 39
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)


        sh_temp = self.sy_eff_2
        # load_reg SY-2-8A
        ind1 = (144,3)
        des = (181,4)
        row_c = 29
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)


        sh_temp = self.nt_eff_0p19
        # load_reg NT-0-1.9A
        ind1 = (184,3)
        des = (142,16)
        row_c = 39
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)


        sh_temp = self.nt_eff_2
        # load_reg NT-2-8A
        ind1 = (144,3)
        des = (181,16)
        row_c = 29
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.Buck_eff_example,all0=0)

        # ===after finished the table, copy the sheet to related place and re-name

        # copy the result in front of the first sheet (SY_0p19)
        self.sh_comp = self.sh_comp.copy(self.sh_sy_eff_0p19)
        temp_name = self.sh_comp.range('C4').value
        self.sh_comp.name = str(temp_name) + '_EFF_comp'



        pass

    def LDO_load_regulation(self):

        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sh_comp = self.wb.sheets("temp")
            self.sh_sy_eff_0p19 = self.wb.sheets("temp")

        self.sh_comp = self.wb.sheets(self.LDO_reg_example)

        # 0p19 => LOD load regulation
        # 2 => Buck on load regulation
        self.buck_sh_name = self.sh_comp.range('C4').value
        self.sy_eff_0p19 = self.sh_comp.range('C5').value
        self.sy_eff_2 = self.sh_comp.range('C6').value
        self.nt_eff_0p19 = self.sh_comp.range('C8').value
        self.nt_eff_2 = self.sh_comp.range('C9').value

        # this is the index sheet for copy summary sheet to front of first sheet
        self.sh_sy_eff_0p19 = self.wb.sheets(self.sy_eff_0p19)


        sh_temp = self.sy_eff_0p19
        # eff SY-LDO load regulation
        ind1 = (224,3)
        des = (43,4)
        row_c = 49
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.LDO_reg_example,all0=0)


        sh_temp = self.sy_eff_2
        # eff SY-Buck on
        ind1 = (224,3)
        des = (142,4)
        row_c = 49
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.LDO_reg_example,all0=0)


        sh_temp = self.nt_eff_0p19
        # eff NT-LDO load regulation
        ind1 = (224,3)
        des = (43,16)
        row_c = 49
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.LDO_reg_example,all0=0)


        sh_temp = self.nt_eff_2
        # eff NT-Buck on
        ind1 = (224,3)
        des = (142,16)
        row_c = 49
        col_c = 7
        self.table_copy(ind_1=ind1, row_count=row_c, col_count=col_c, ind_dest0=des, sheet_sor0=sh_temp, sheet_des0=self.LDO_reg_example,all0=0)

        # ===after finished the table, copy the sheet to related place and re-name

        # copy the result in front of the first sheet (SY_0p19)
        self.sh_comp = self.sh_comp.copy(self.sh_sy_eff_0p19)
        temp_name = self.sh_comp.range('C4').value
        self.sh_comp.name = str(temp_name) + '_LDO_reg_comp'

        pass

    def table_comparison(self,ind_1=(1, 1),ind_2=(3, 3),row_count=3,col_count=3,ind_dest0=(5, 5),sheet_sor0="",sheet_des0="",diff=1):
        """
        copy table and generate comparison table
        diff = 1 => add the difference comparison
        """
        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sheet_sor = self.wb.sheets("temp")
            self.sheet_des = self.wb.sheets("temp")

        ind_1_end = (ind_1[0] + row_count, ind_1[1] + col_count)
        ind_2_end = (ind_2[0] + row_count, ind_2[1] + col_count)

        self.sor_des_sheet_sel(sheet_sor1=sheet_sor0, sheet_des1=sheet_des0)

        # sor_range().copy(res_range)

        # shift destination index
        ind_dest1 = (ind_dest0[0], ind_dest0[1] + col_count + self.space)
        ind_dest2 = (ind_dest0[0], ind_dest1[1] + col_count + self.space)

        sor_range = self.sheet_sor.range(ind_1, ind_1_end)
        res_range = self.sheet_des.range(ind_dest0)

        # copy table 1
        sor_range.copy(res_range)

        sor_range = self.sheet_sor.range(ind_2, ind_2_end)
        res_range = self.sheet_des.range(ind_dest1)

        # cpoy table 2
        sor_range.copy(res_range)

        sor_range = self.sheet_sor.range(ind_2, ind_2_end)
        res_range = self.sheet_des.range(ind_dest2)

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
                    self.sheet_des.range(ind_r2, ind_c2).value = (
                        self.sheet_des.range(ind_r0, ind_c0).value - self.sheet_des.range(ind_r1, ind_c1).value
                    )

                    x_col = x_col + 1
                    pass

                x_row = x_row + 1
                pass

        pass

    def table_copy(self,ind_1=(1, 1),row_count=3,col_count=3,ind_dest0=(5, 5),sheet_sor0="",sheet_des0="",all0=1):
        '''
        just copy one table, not for comparison \n
        sheet_sor0, sheet_des0 should be string \n
        if all0 == 0 only copy item, not the index row and column \n
        '''
        prog_fake = 0
        if prog_fake == 1:
            self.wb = xw.Book()
            self.sheet_sor = self.wb.sheets("temp")
            self.sheet_des = self.wb.sheets("temp")

        if all0 == 0 :
        # shift and not copy the index row and column
            ind_1 = (ind_1[0] + 1, ind_1[1] + 1)
            # ind_1[0] = ind_1[0] + 1
            # ind_1[1] = ind_1[1] + 1
            row_count = row_count - 1
            col_count = col_count - 1



        ind_1_end = (ind_1[0] + row_count, ind_1[1] + col_count)

        self.sor_des_sheet_sel(sheet_sor1=sheet_sor0,sheet_des1=sheet_des0)

        # sor_range().copy(res_range)

        sor_range = self.sheet_sor.range(ind_1, ind_1_end)
        res_range = self.sheet_des.range(ind_dest0)

        # copy table 1
        # sor_range.copy(res_range)
        res_range.value = sor_range.value


        pass

    def table_title(self):
        '''
        use this program to copy the title of table
        '''

        pass

    def sor_des_sheet_sel(self, sheet_sor1='', sheet_des1=''):
        '''
        select the sheet of sor and des, sor is must have and des is option \n
        default option => sor \n
        both input will be name, not sheet_obj
        '''

        if sheet_sor1 != '':
            self.sheet_sor = self.wb.sheets(str(sheet_sor1))
            # if no sor0 input, set to default, may be error
            pass

        if sheet_des1 != '':
            self.sheet_des = self.wb.sheets(str(sheet_des1))
            pass
        else:
            # if no des input, use the same sheet
            self.sheet_des = self.sheet_sor


        pass


    pass




if __name__ == "__main__":
    #  the testing code for this file object

    operation_index = 2
    if operation_index == 0 :
        test_index = 2
        trace = 0
        file_name0="test.xlsx"
    else:
        test_index = 100
        trace = 1
        file_name = 'c:\\py_gary\\test_excel\\NT50970NAB_verification.xlsm'

    rep = report_arragement(file_name0=file_name,full_trace0=trace)

    if operation_index == 1:

        rep.Buck_eff_load_regulation()

        pass
    elif operation_index == 2 :

        rep.LDO_load_regulation()



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
