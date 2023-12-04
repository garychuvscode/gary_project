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
    def __init__(self,excel0=0):
        """
        this class used to process some regular copy, paste and plot of report \n
        file input need to be .xlsm
        """

        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst
            import parameter_load_obj as par

            # add the libirary from Geroge
            import Scope_LE6100A as sco

            # initial the object and set to simulation mode
            pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, "off")
            pwr0.sim_inst = 0
            # initial the object and set to simulation mode
            met_v0 = inst.Met_34460(0.0001, 30, 0.000001, 2.5, 21)
            met_v0.sim_inst = 0
            loader_0 = inst.chroma_63600(1, 7, "CCL")
            loader_0.sim_inst = 0
            # mcu is also config as simulation mode
            mcu0 = mcu.MCU_control(0, 3)
            # using the main control book as default
            excel0 = par.excel_parameter("obj_main")
            src0 = inst.Keth_2440(0, 0, 24, "off", "CURR", 15)
            src0.sim_inst = 0
            met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
            met_i0.sim_inst = 0
            chamber0 = inst.chamber_su242(25, 10, "off", -45, 180, 0)
            chamber0.sim_inst = 0
            scope0 = sco.Scope_LE6100A("GPIB: 15", 0, 0)
            # ======== only for object programming

        self.full_trace = ''
        self.excel_ini = excel0
        self.file_name = ''
        self.wb = 0
        self.default_path = 'c:\\py_gary\\test_excel\\'




        # name of sheet for example operation
        self.Buck_eff_example = "Eff_comp_ex_C"
        self.LDO_reg_example = "regulation_comp_LDO_ex_C"
        self.Buck_eff_example_B = "Eff_comp_ex_AB"
        self.LDO_reg_example_B = "regulation_comp_LDO_ex_AB"
        # this sheet is used to grab data from the verification report raw data
        self.raw_sheet = "C_raw_NAB"

        # this is not sheet name, already sheet object
        self.sheet_sor = ""
        self.sheet_des = ""

        # name of source sheet
        self.source_sheet = 0
        self.destination_sheet = 0

        # spece between each table (column + space)
        self.space = 4

        self.default_tag_name = f'<virtual> #1'

        # 231011 fixed temp result trace
        # buffer between report manager and GPL_V5
        self.temp_res_bufffer = 'c:\\py_gary\\test_excel\\GPL_V5_RPC_temp.xlsx'
        # window and other reference buffer:
        self.grace_trace = 'c:\\py_gary\\test_excel\\grace_trace.xlsx'



        pass
    def open_book(self, full_trace0=0, file_name0=0, check_ref_sh0=1):

        if full_trace0 == 0 :

            # input file name of report, which already opened
            self.file_name = str(file_name0)
            self.wb = self.excel_ini.app_org.books(self.file_name)
            # self.wb = xw.books(self.file_name)
            print(f"select {self.file_name} as report file")
            self.full_trace = ''
            pass
        else:
            # file saving is based on the trace content
            self.wb = self.excel_ini.app_org.books.open(file_name0)
            # self.wb = xw.Book(file_name0)
            self.full_trace = file_name0
            pass

        # check the opened book have ref_sh or not
        if check_ref_sh0 == 1 :
            try :
                self.ref_sh = self.wb.sheets('ref_sh')
                print(f'reference sheet already exist, start the arragement')
                pass

            except :
                print(f'reference sheet not exist, add a new one and assigned')
                self.ref_sh = self.wb.sheets.add(name='ref_sh')

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

        # # this one only need for the old solutions before 231012
        # # this is the index sheet for copy summary sheet to front of first sheet


        # copy the sheet to result book first
        self.find_sh_all_books_copy_res(sheet_name=self.sy_eff_0p19)
        self.find_sh_all_books_copy_res(sheet_name=self.sy_eff_2)
        self.find_sh_all_books_copy_res(sheet_name=self.nt_eff_0p19)
        self.find_sh_all_books_copy_res(sheet_name=self.nt_eff_2)

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

        # this will copy the format
        # sor_range.copy(res_range)

        # this will only copy the value
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

        if sheet_sor1 != '' :
            # # new version: search the sheet from global, and copy to result book (self.wb)
            # self.sheet_sor = self.find_sh_all_books_copy_res(sheet_name=str(sheet_sor1))
            # version before 231012
            self.sheet_sor = self.wb.sheets(str(sheet_sor1))
            # if no sor0 input, set to default, may be error
            pass

        if sheet_des1 != '':
            # # new version: search the sheet from global, and copy to result book (self.wb)
            # self.sheet_des = self.find_sh_all_books_copy_res(sheet_name=str(sheet_des1), copy0=0)
            # version before 231012
            self.sheet_des = self.wb.sheets(str(sheet_des1))
            pass
        else:
            # if no des input, use the same sheet
            self.sheet_des = self.sheet_sor


        pass

    def report_temp_ini(self, file_name0=0):
        '''
        make the temp file status back to intial

        GPL_V5_RPC_temp.xlsx => only leave the ref_sh, delete other sheet
        grace_trace.xlsx => delete other sheet and copy ref_sh to replace the sheet

        fixed trace:
        'c:\\py_gary\\test_excel\\file_name0'
        '''

        self.clr_to_ref_sh(file_name0)

        if file_name0 == "grace_trace.xlsx" :
            # also need to copy ref_sh and create new sheet beofre close
            sh_ref = self.ini_wb.sheets('ref_sh')
            sh_temp = sh_ref.copy(sh_ref)
            sh_temp.name = 'file_trace_ref'
            pass

        self.ini_wb.save()
        self.ini_wb.close()

    pass

    def clr_to_ref_sh(self, file_name0=0):
        '''
        file_name0 need to be full file name: test.xlsx or test.xlsm
        '''
        full_trace = self.default_path + file_name0

        try:
            # 打开指定的 Excel 文件
            self.ini_wb = self.excel_ini.app_org.books.open(full_trace)

            # 遍历所有工作表，除了名为 "sh_ref" 的工作表之外，都删除
            for sheet in self.ini_wb.sheets:
                if sheet.name != "ref_sh":
                    sheet.delete()

            # 保存对文件的更改
            self.ini_wb.save()

            # # 关闭文件和 Excel 应用程序
            # wb.close()
            # app.quit()

            print(f"已删除{file_name0}所有工作表，只保留 'ref_sh' 工作表。")

        except Exception as e:
            print(f"发生错误：{str(e)}")

        pass

    def find_sh_all_books_copy_res(self, sheet_name='', copy0=1):
        '''
        mode0 is set to 7 in default, it will send 'other settings in tag'
        otherwise, it will be the tage support summary sheet
        decide to copy the sheet or not
        '''
        # 连接到 Excel 应用程序
        # app = xw.App(visible=False)  # 如果不需要可见 Excel，请设置 visible=False
        app2 = xw.apps
        finded = 0

        print(app2)
        for app in app2:
            print(f'now is app{app}')
            workbooks = app.books
            # 遍历所有已打开的工作簿
            for workbook in workbooks:
                print(f'now is wb{workbook}')
                # 遍历当前工作簿中的所有工作表
                for sheet in workbook.sheets:
                    print(f'now is wb{sheet}')
                    if sheet.name == sheet_name:
                        # 找到匹配的工作表
                        print(f"在app{app}工作簿 '{workbook.name}' 中找到匹配的工作表：{sheet_name}")
                        # 在这里可以执行其他操作，如读取或修改工作表内容
                        '''
                        copy this sheet to the result book, it should be able to find the
                        result book and reference sheet for copying index during the program operation
                        '''
                        if copy0 == 1 :
                            # 231103: add comparison, if in same book, no need to copy
                            book_check = self.sh_comp.book
                            if book_check != workbook :
                                sheet = sheet.copy(self.sh_comp)
                                pass
                            pass
                        finded = 1
                        break
                    if finded == 1 :
                        break
                    pass
                if finded == 1 :
                    break
                pass
            if finded == 1 :
                break
            pass
        # return the finded sheet
        return sheet

    def buck_eff_sum_gen_full(self, mode0=0, setting_sel0='virtual', L_H0=1):
        '''
        auto generation for full report of buck eff and regulation
        default tag name : f'<{tag_set}> #1'
        mode 0 => all, 1-LDO(L), 2-LDO_BK_on(LBO), 3-Buck_normal(BN), 4-Buck_usm(BU), 5-Buck_normal_H, 6-Buck_usm_H
        231012 => since there are comparison needed for summary sheet, no need this function now
        '''
        self.default_tag_name = f'<{setting_sel0}> #1'
        # watch out the index change due to process of exp and report is different
        # self.mode_index = [ '_L', '_LBO', '_BN', '_BU', '_BN_H', '_BU_H' ]
        self.mode_index = [ '_L', '_LBO', '_BN', '_BN_H', '_BU', '_BU_H' ]

        app = self.excel_ini.app_org
        temp_wb = app.books.open(self.temp_res_bufffer)

        if mode0 != 0 :
            # run a full result
            x_mode = 0

            if L_H0 == 0 :
                c_mode = 4
            else:
                c_mode = 6

            while x_mode < c_mode :

                # mode x_mode
                tag_set = setting_sel0 + self.mode_index[x_mode]
                temp_sh = temp_wb.sheets(tag_set)

                # copy to result sheet
                temp_sh = temp_sh.copy(self.excel_ini.sh_ref)

                if x_mode < 2 :
                    # itmes for LDO regulation
                    pass
                elif x_mode < 4 :
                    # buck normal mode operation report summary
                    pass
                else :
                    # buck USM mode operation report summary
                    pass

                print(f'all in one result generationmpde = {mode0}')
                x_mode = x_mode + 1
                pass
            print(f'all in one result generation done')


            # LDO regulation

            # Buck-normal regulation

            # Buck-USM regulation


        # LDO regulation

        # Buck-normal regulation

        # Buck-USM regulation


        pass

    def buck_eff_sum_gen_single(self, setting_sel0='virtual', L_H0=1):
        '''
        choose by using message box:
        1. C or B version selection
        2. high and low sheet assigned (choose active and press enter)

        '''
        pass

    def new_file_from_temp(self, file_name_org0='GPL_V5_RPC_temp', file_name_dest0='', file_name_extra0='_buck_reg_mix'):
        '''
        this function create books with new name, all in testexcel
        default file name: 'GPL_V5_RPC_temp'
        '''

        if file_name_dest0 == '' :
            file_name_dest0 = self.excel_ini.eff_sh_name
        # open and create new file name for the final result
        book_res_n = self.excel_ini.app_org.books.open(self.temp_res_bufffer)
        # for change book's name, need to use save
        new_name = file_name_dest0 + file_name_extra0 +'.xlsx'
        book_res_n.save(path=self.default_path+str(new_name))
        print(f'finished test and create the final file at {new_name}')
        book_res_n.close()


if __name__ == "__main__":
    #  the testing code for this file object

    import sheet_ctrl_main_obj as sh
    import parameter_load_obj as para

    excel_m = para.excel_parameter(str(sh.file_setting))

    operation_index = 1
    if operation_index == 0 :
        test_index = 3
        trace = 0
        file_name="test.xlsx"
    else:
        # testing index to 100, no need testing
        test_index = 100
        trace = 1
        file_name = 'c:\\py_gary\\test_excel\\NT50970NAB_verification.xlsm'

    rep = report_arragement(excel0=excel_m)


    if operation_index == 1:
        rep.open_book(file_name0=file_name,full_trace0=trace, check_ref_sh0=0)

        rep.Buck_eff_load_regulation()

        pass
    elif operation_index == 2 :
        rep.open_book(file_name0=file_name,full_trace0=trace, check_ref_sh0=0)

        rep.LDO_load_regulation()



    if test_index == 0 :
        rep.open_book(file_name0=file_name,full_trace0=trace, check_ref_sh0=0)

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
        rep.open_book(file_name0=file_name,full_trace0=trace, check_ref_sh0=0)
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
        rep.open_book(file_name0=file_name,full_trace0=trace, check_ref_sh0=0)

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

    elif test_index == 3:
        # testing for clr temp file:

        # GPL_V5_RPC_temp.xlsx => only leave the ref_sh, delete other sheet
        # grace_trace.xlsx => delete other sheet and copy ref_sh to replace the sheet

        rep.report_temp_ini(file_name0='GPL_V5_RPC_temp.xlsx')
        rep.report_temp_ini(file_name0='grace_trace.xlsx')


    pass
