# 220919 this file is used to generate the format sheet for waveform
# change the py file to a object, so it can be used to generate many sheet
# in a same file and overcome one report in the end
# how to choose if to separate the file or not?
# maybe don't think about to separate the file will be easier
# first version not to separate file


# import the excel control package
import xlwings as xw
# numpy is used for matrix operation, check more from google
import numpy as np
# excel parameter and settings
import parameter_load_obj as par


class format_gen:

    # this class is used to measure IQ from the DUT, based on the I/O setting and different Vin
    # measure the IQ

    def __init__(self, excel0, ctrl_sheet_name0):

        # # ======== only for object programming
        # # testing used temp instrument
        # # need to become comment when the OBJ is finished
        # import mcu_obj as mcu
        # import inst_pkg_d as inst
        # # initial the object and set to simulation mode
        # pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
        # pwr0.sim_inst = 0
        # # initial the object and set to simulation mode
        # met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
        # met_v0.sim_inst = 0
        # loader_0 = inst.chroma_63600(1, 7, 'CCL')
        # loader_0.sim_inst = 0
        # # mcu is also config as simulation mode
        # mcu0 = mcu.MCU_control(0, 3)
        # # using the main control book as default
        # excel0 = par.excel_parameter('obj_main')
        # src0 = inst.Keth_2440(0, 0, 24, 'off', 'CURR', 15)
        # src0.sim_inst = 0
        # met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
        # met_i0.sim_inst = 0
        # chamber0 = inst.chamber_su242(25, 10, 'off', -45, 180, 0)
        # chamber0.sim_inst = 0
        # # ======== only for object programming

        # assign the input information to object variable
        self.excel_ini = excel0

        # assign the related sheet of each format gen
        self.ctrl_sheet_name = ctrl_sheet_name0
        self.excel_ini.sh_format_gen = self.excel_ini.wb.sheets(
            str(self.ctrl_sheet_name))
        self.sh_format_gen = self.excel_ini.wb.sheets(
            str(self.ctrl_sheet_name))

        # use another experiment object to send the name in, and
        # related object will adjust the related instrument for the experiment
        # this object only for the testing result format generation
        self.sh_ref_table = self.excel_ini.wb.sheets('table')

        # loading the control values
        self.c_row_item = self.sh_format_gen.range('C31').value
        self.c_column_item = self.sh_format_gen.range('C32').value
        self.c_data_mea = self.sh_format_gen.range('C33').value

        self.item_str = self.sh_format_gen.range('C28').value
        self.row_str = self.sh_format_gen.range('C29').value
        self.col_str = self.sh_format_gen.range('C30').value
        self.extra_str = self.sh_format_gen.range('C34').value
        self.color_default = self.sh_format_gen.range('C37').color
        self.color_target = self.sh_format_gen.range('C38').color

        self.target_width = self.sh_format_gen.range('I2').column_width
        self.target_heigh = self.sh_format_gen.range('I2').row_height
        self.default_width = self.sh_format_gen.range('J5').column_width
        self.default_heigh = self.sh_format_gen.range('J5').row_height
        # default height and width is used to prevent shape change of the table
        # target height and width is used to save the waveform capture from scope

        # start to adjust the the format based on the input settings

        self.new_sheet_name = str(self.sh_format_gen.range('C35').value)

        pass

    def run_format_gen(self):

        # start to adjust the the format based on the input settings

        # before changin the format, adjust the color
        y_dim = 0
        c_y_dim = 3 * self.c_column_item
        # y dimension 3*c_column_item, because not adding c_data_mea yet
        # knowing the amount is enough
        while y_dim < c_y_dim:
            # need to check every cell in the effective operating range
            # from the sheet setting in CTRL sheet

            x_dim = 0
            c_x_dim = self.c_row_item + 1
            while x_dim < c_x_dim:
                color_temp = self.sh_ref_table.range(
                    (4 + y_dim, 1 + x_dim)).color

                if color_temp == self.color_default:
                    self.sh_ref_table.range(
                        (4 + y_dim, 1 + x_dim)).color = self.color_target
                    # change the needed place with default color to the target color

                x_dim = x_dim + 1

            y_dim = y_dim + 1

        # before insert the row, add the content into realted row and column
        # use the same definition but different action
        y_dim = 0
        c_y_dim = self.c_column_item
        # y dimension 3*c_column_item, because not adding c_data_mea yet
        # knowing the amount is enough

        # filter the error of extra_str = none (error handling for the no input cells)
        extra_str = self.extra_str
        item_str = self.item_str
        row_str = self.row_str
        col_str = self.col_str

        if extra_str == None:
            extra_str = ''
        if item_str == None:
            item_str = ''
        if row_str == None:
            row_str = ''
        if col_str == None:
            col_str = ''

        # x_dim, y_dim are the dimension counter for modifing the table
        while y_dim < c_y_dim:
            # need to check every cell in the effective operating range
            # from the sheet setting in CTRL sheet
            str_temp = self.sh_format_gen.range((43 + y_dim, 2)).value
            excel_temp = col_str + '\n' + \
                str(str_temp) + item_str + '\n' + extra_str
            self.sh_ref_table.range((4 + 1 + y_dim * 3, 1)).value = excel_temp
            self.sh_ref_table.range(
                (4 + 1 + y_dim * 3, 1)).row_height = self.target_heigh
            # height need to change when modifing the column cells

            x_dim = 0
            c_x_dim = self.c_row_item
            while x_dim < c_x_dim:

                str_temp = self.sh_format_gen.range((43 + x_dim, 1)).value
                excel_temp = row_str + str(str_temp)
                self.sh_ref_table.range(
                    (4 + y_dim * 3, 1 + 1 + x_dim)).value = excel_temp
                self.sh_ref_table.range((4 + y_dim * 3, 1 + 1 + x_dim)
                                        ).column_width = self.target_width
                # width need to change when modifing the row cells

                x_dim = x_dim + 1

            y_dim = y_dim + 1

        # add the new row to the related position
        x_column_item = 0
        while x_column_item < self.c_column_item:
            # each column item need to have related data result row

            x_data_mea = 0
            while x_data_mea < self.c_data_mea:
                # will need to insert the row and assign the row index at the same time
                # first to insert the related row with new color
                if x_data_mea > 0:
                    self.sh_ref_table.api.Rows(
                        6 + (2 + self.c_data_mea) * x_column_item).Insert()

                # then assign the related data name for related row ( in reverse order )
                excel_temp = self.sh_format_gen.range(
                    (43 + self.c_data_mea - x_data_mea - 1, 3)).value
                self.sh_ref_table.range(
                    (6 + (2 + self.c_data_mea) * x_column_item, 1)).value = excel_temp
                # keep the added row in the default high, not change due to insert
                self.sh_ref_table.range(
                    (6 + (2 + self.c_data_mea) * x_column_item, 1)).row_height = self.default_heigh

                x_data_mea = x_data_mea + 1
                # testing of insert row into related position
                # self.sh_ref_table.api.Rows(6).Insert()

            x_column_item = x_column_item + 1

        self.excel_ini.excel_save()
        # wb_res.save(result_book_trace)
        # save the result after program is finished

        pass



    def sheet_gen(self):
        # this function is a must have function to generate the related excel for this verification item
        # this sub must include:
        # 2. generate the result sheet in the result book, and setup the format
        # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
        # 4. not a new file but an add on sheet to the result workbook

        # copy the rsult sheet to result book
        self.excel_ini.sh_format_gen.copy(self.excel_ini.sh_ref)
        # assign the sheet to result book
        self.excel_ini.sh_format_gen = self.excel_ini.wb_res.sheets(
            str(self.ctrl_sheet_name))
        # copy the rsult sheet to result book
        self.excel_ini.sh_ref_table.copy(self.excel_ini.sh_ref)
        # assign the sheet to result book
        self.excel_ini.sh_ref_table = self.excel_ini.wb_res.sheets('table')

        # change the sheet name after finished and save into the excel object
        self.excel_ini.sh_ref_table.name = str(self.new_sheet_name)
        self.sh_ref_table = self.excel_ini.sh_ref_table

        # # 220914 move to build file to prevent error of delete and end of file
        # # copy the result sheet to result book
        # self.excel_ini.sh_raw_out.copy(self.excel_ini.sh_ref)
        # # assign the sheet to result book
        # self.excel_ini.sh_raw_out = self.excel_ini.wb_res.sheets('raw_out')

        # # copy the sheets to new book
        # # for the new sheet generation, located in sheet_gen
        # self.excel_s.sh_main.copy(self.sh_ref_condition)
        # self.sh_result.copy(self.sh_ref)

        pass

    def table_return(self):
            # need to recover this sheet: self.excel_ini.sh_ref_table
            self.excel_ini.sh_ref_table = self.excel_ini.wb.sheets('table')

            pass


if __name__ == '__main__':
    #  the testing code for this file object

    # ======== only for object programming
    # testing used temp instrument
    # need to become comment when the OBJ is finished
    import mcu_obj as mcu
    import inst_pkg_d as inst
    # initial the object and set to simulation mode
    pwr_t = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
    pwr_t.sim_inst = 0
    # initial the object and set to simulation mode
    met_v_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_v_t.sim_inst = 0
    load_t = inst.chroma_63600(1, 7, 'CCL')
    load_t.sim_inst = 0
    met_i_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_i_t.sim_inst = 0
    src_t = inst.Keth_2440(0, 0, 24, 'off', 'CURR', 15)
    src_t.sim_inst = 0
    chamber_t = inst.chamber_su242(25, 10, 'off', -45, 180, 0)
    chamber_t.sim_inst = 0
    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu_t = mcu.MCU_control(0, 3)

    # for the single test, need to open obj_main first,
    # the real situation is: sheet_ctrl_main_obj will start obj_main first
    # so the file will be open before new excel object benn define

    # using the main control book as default
    excel_t = par.excel_parameter('obj_main')
    # ======== only for object programming

    # open the result book for saving result
    excel_t.open_result_book()

    # change simulation mode delay (in second)
    excel_t.sim_mode_delay(0.02, 0.01)
    inst.wait_time = 0.01
    inst.wait_samll = 0.01

    # and the different verification method can be call below

    # create one file
    format_gen_1 = format_gen(excel_t, 'CTRL_sh_ex')

    # generate(or copy) the needed sheet to the result book
    format_gen_1.sheet_gen()

    # start the testing
    format_gen_1.run_format_gen()
    # return the table after data input finished
    format_gen_1.table_return()

    # create one file
    format_gen_2 = format_gen(excel_t, 'CTRL_sh_line')

    # generate(or copy) the needed sheet to the result book
    format_gen_2.sheet_gen()

    # start the testing
    format_gen_2.run_format_gen()
    format_gen_2.table_return()

    # create one file
    format_gen_3 = format_gen(excel_t, 'CTRL_sh_load')

    # generate(or copy) the needed sheet to the result book
    format_gen_3.sheet_gen()

    # start the testing
    format_gen_3.run_format_gen()
    format_gen_3.table_return()

    # remember that this is only call by main, not by  object
    excel_t.end_of_file(0)

    print('end of the SWIRE scan object testing program')
