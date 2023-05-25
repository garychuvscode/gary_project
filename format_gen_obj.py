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

    def __init__(self, excel0):
        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst

            # initial the object and set to simulation mode
            pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, "off")
            pwr0.sim_inst = 0
            # initial the object and set to simulation mode
            met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
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
            # ======== only for object programming

        # assign the input information to object variable
        self.excel_ini = excel0

        # # 220926 change the call format
        # # assign the related sheet of each format gen
        # self.ctrl_sheet_name = ctrl_sheet_name0
        # self.excel_ini.sh_format_gen = self.excel_ini.wb.sheets(
        #     str(self.ctrl_sheet_name))
        # self.sh_format_gen = self.excel_ini.wb.sheets(
        #     str(self.ctrl_sheet_name))

        # use another experiment object to send the name in, and
        # related object will adjust the related instrument for the experiment
        # this object only for the testing result format generation
        self.sh_ref_table = self.excel_ini.wb.sheets("table")

        # indicator for the sheet name input check 0 => no correct name, 1 is ok
        self.sheet_name_ready = 0

        pass

    def set_sheet_name(self, ctrl_sheet_name0):
        # 221212: add the table return inisde the object
        # return the table before next use
        self.table_return()

        # assign the related sheet of each format gen
        self.ctrl_sheet_name = ctrl_sheet_name0

        # sh_format_gen is the sheet can be access from other object
        # load the setting value for instrument
        self.excel_ini.sh_format_gen = self.excel_ini.wb.sheets(
            str(self.ctrl_sheet_name)
        )
        self.sh_format_gen = self.excel_ini.wb.sheets(str(self.ctrl_sheet_name))

        # also include the new sheet setting from each different sheet

        # loading the control values
        # 220926: index of counter need to passed to the excel, so other object or instrument
        # is able to reference

        self.c_row_item = self.sh_format_gen.range("C31").value
        self.c_column_item = self.sh_format_gen.range("C32").value
        # c_data_mea is data count
        self.c_data_mea = self.sh_format_gen.range("C33").value
        self.c_ctrl_var1 = self.sh_format_gen.range("D40").value
        self.c_ctrl_var2 = self.sh_format_gen.range("E40").value
        self.c_ctrl_var4 = self.sh_format_gen.range("G40").value

        # also need to assign the variable to the excel obj
        self.excel_ini.c_row_item = self.c_row_item
        self.excel_ini.c_column_item = self.c_column_item
        self.excel_ini.c_data_mea = self.c_data_mea
        self.excel_ini.c_ctrl_var1 = self.c_ctrl_var1
        self.excel_ini.c_ctrl_var2 = self.c_ctrl_var2
        self.excel_ini.c_ctrl_var4 = self.c_ctrl_var4

        self.item_str = self.sh_format_gen.range("C28").value
        self.row_str = self.sh_format_gen.range("C29").value
        self.col_str = self.sh_format_gen.range("C30").value
        self.extra_str = self.sh_format_gen.range("C34").value
        self.color_default = self.sh_format_gen.range("C37").color
        self.color_target = self.sh_format_gen.range("C38").color

        self.target_width = self.sh_format_gen.range("I2").column_width
        self.target_height = self.sh_format_gen.range("I2").row_height
        self.default_width = self.sh_format_gen.range("J5").column_width
        self.default_height = self.sh_format_gen.range("J5").row_height
        # default height and width is used to prevent shape change of the table
        # target height and width is used to save the waveform capture from scope

        # start to adjust the the format based on the input settings

        self.new_sheet_name = str(self.sh_format_gen.range("C35").value)

        print("sheet name ready")
        self.sheet_name_ready = 1
        self.sheet_gen()
        self.run_format_gen()

        pass

    def run_format_gen(self):
        if self.sheet_name_ready == 0:
            print("no proper sheet name set yet, need to set_sheet_name first")
            pass

        else:
            self.reload_waveform_dimension()
            # start to adjust the the format based on the input settings

            x_sheets = 0
            # add one more loop for the multi pulse or i2c command generation
            while x_sheets < self.c_sheets:
                self.sh_ref_table = self.excel_ini.ref_table_list[x_sheets]

                # before changing the format, adjust the color
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
                            (4 + y_dim, 1 + x_dim)
                        ).color

                        if color_temp == self.color_default:
                            self.sh_ref_table.range(
                                (4 + y_dim, 1 + x_dim)
                            ).color = self.color_target
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
                    extra_str = ""
                if item_str == None:
                    item_str = ""
                if row_str == None:
                    row_str = ""
                if col_str == None:
                    col_str = ""

                # x_dim, y_dim are the dimension counter for modifing the table
                while y_dim < c_y_dim:
                    # need to check every cell in the effective operating range
                    # from the sheet setting in CTRL sheet
                    str_temp = self.sh_format_gen.range((43 + y_dim, 2)).value

                    # 221102 insert for summary table
                    # self.excel_ini.sum_table_gen(self.excel_ini.summary_start_x, self.excel_ini.summary_start_y, 0, y_dim + 1, content = str_temp)
                    self.summary_fo_gen(
                        0,
                        y_dim + 1,
                        self.c_row_item,
                        self.c_column_item,
                        x_sheets,
                        self.c_data_mea,
                        str_temp,
                    )

                    excel_temp = (
                        col_str + "\n" + str(str_temp) + item_str + "\n" + extra_str
                    )
                    self.sh_ref_table.range((4 + 1 + y_dim * 3, 1)).value = excel_temp
                    self.sh_ref_table.range(
                        (4 + 1 + y_dim * 3, 1)
                    ).row_height = self.target_height
                    # height need to change when modifing the column cells

                    x_dim = 0
                    c_x_dim = self.c_row_item
                    while x_dim < c_x_dim:
                        str_temp = self.sh_format_gen.range((43 + x_dim, 1)).value
                        # 221102 insert for summary table
                        # self.excel_ini.sum_table_gen(self.excel_ini.summary_start_x, self.excel_ini.summary_start_y, x_dim + 1, 0, content = str_temp)
                        self.summary_fo_gen(
                            x_dim + 1,
                            0,
                            self.c_row_item,
                            self.c_column_item,
                            x_sheets,
                            self.c_data_mea,
                            str_temp,
                        )

                        excel_temp = row_str + str(str_temp)
                        self.sh_ref_table.range(
                            (4 + y_dim * 3, 1 + 1 + x_dim)
                        ).value = excel_temp
                        self.sh_ref_table.range(
                            (4 + y_dim * 3, 1 + 1 + x_dim)
                        ).column_width = self.target_width
                        # width need to change when modifing the row cells

                        x_dim = x_dim + 1
                        pass

                    y_dim = y_dim + 1
                    pass

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
                                6 + (2 + self.c_data_mea) * x_column_item
                            ).Insert()

                        # then assign the related data name for related row ( in reverse order )
                        excel_temp = self.sh_format_gen.range(
                            (43 + self.c_data_mea - x_data_mea - 1, 3)
                        ).value

                        if x_column_item == 0:
                            # 221102 insert for summary table  (measurement index and the pulse or I2C setting)
                            self.excel_ini.sum_table_gen(
                                self.excel_ini.summary_start_x,
                                self.excel_ini.summary_start_y,
                                0
                                + (self.c_data_mea - x_data_mea - 1)
                                * (c_x_dim + self.excel_ini.summary_gap),
                                0 + x_sheets * (c_y_dim + self.excel_ini.summary_gap),
                                content=excel_temp,
                            )

                            if self.i2c_en == 0:
                                # pulse command
                                data1 = self.sh_format_gen.range(
                                    (43 + x_sheets, 5)
                                ).value
                                data2 = self.sh_format_gen.range(
                                    (43 + x_sheets, 6)
                                ).value
                                condition_str = (
                                    f"pulse cmd: {int(data1)} and {int(data2)}"
                                )
                            else:
                                condition_str = f"I2C cmd: group {x_sheets + 1}"

                            #  shift one cell above the measurement content
                            self.excel_ini.sum_table_gen(
                                self.excel_ini.summary_start_x,
                                self.excel_ini.summary_start_y,
                                0
                                + (self.c_data_mea - x_data_mea - 1)
                                * (c_x_dim + self.excel_ini.summary_gap),
                                0
                                + x_sheets * (c_y_dim + self.excel_ini.summary_gap)
                                - 1,
                                content=condition_str,
                            )

                        self.sh_ref_table.range(
                            (6 + (2 + self.c_data_mea) * x_column_item, 1)
                        ).value = excel_temp
                        # keep the added row in the default high, not change due to insert
                        self.sh_ref_table.range(
                            (6 + (2 + self.c_data_mea) * x_column_item, 1)
                        ).row_height = self.default_height

                        x_data_mea = x_data_mea + 1
                        # testing of insert row into related position
                        # self.sh_ref_table.api.Rows(6).Insert()
                        pass

                    x_column_item = x_column_item + 1
                    pass
                x_sheets = x_sheets + 1
                pass

            self.excel_ini.excel_save()
            # wb_res.save(result_book_trace)
            # save the result after program is finished
            pass

        pass

    def summary_fo_gen(
        self, x_ind, y_ind, c_x_item, c_y_item, x_sheet_s, c_mea=1, content=0
    ):
        # x_sheets = 0
        # while x_sheets < self.c_sheets:
        for i in range(0, int(c_mea)):
            # 221102 insert for summary table
            self.excel_ini.sum_table_gen(
                self.excel_ini.summary_start_x,
                self.excel_ini.summary_start_y,
                x_ind + i * (c_x_item + self.excel_ini.summary_gap),
                y_ind + x_sheet_s * (c_y_item + self.excel_ini.summary_gap),
                content,
            )

            pass

            # x_sheets = x_sheets + 1

        pass

    def sheet_gen(self):
        if self.sheet_name_ready == 0:
            print("no proper sheet name set yet, need to set_sheet_name first")
            pass
        else:
            # this function is a must have function to generate the related excel for this verification item
            # this sub must include:
            # 2. generate the result sheet in the result book, and setup the format
            # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
            # 4. not a new file but an add on sheet to the result workbook

            # # copy the result sheet to result book
            # self.excel_ini.sh_format_gen.copy(self.excel_ini.sh_ref)
            # # assign the sheet to result book
            # self.excel_ini.sh_format_gen = self.excel_ini.wb_res.sheets(
            #     str(self.ctrl_sheet_name))
            # 221209: since .copy will return the cpoied sheet, just assign, no need for name
            self.excel_ini.sh_format_gen = self.excel_ini.sh_format_gen.copy(
                self.excel_ini.sh_ref
            )

            self.i2c_en = int(self.excel_ini.sh_format_gen.range("B10").value)

            if self.i2c_en == 0:
                self.c_sheets = int(self.excel_ini.sh_format_gen.range("E40").value)
            else:
                self.c_sheets = int(self.excel_ini.sh_format_gen.range("B12").value)

            x_sheets = 0
            c_try = 50
            while x_sheets < self.c_sheets:
                # # copy the result sheet to result book (reference table)
                # self.excel_ini.sh_ref_table.copy(self.excel_ini.sh_ref)
                # # assign the sheet to result book
                # temp_sheet = self.excel_ini.wb_res.sheets(
                #     'table')
                # 221209: since .copy will return the cpoied sheet, just assign, no need for name
                temp_sheet = self.excel_ini.sh_ref_table.copy(self.excel_ini.sh_ref)

                # 221223: to overcome the issue of new file and old file, need to separate the try except function
                if self.excel_ini.keep_last == 0:
                    # to use the new file, keep_last disable

                    # 221218 replace with the try function (re-enable this part 221223)
                    temp_sheet.name = str(self.new_sheet_name + "_" + str(x_sheets))

                else:
                    # 221223: move this part to else, when keep_last == 1, add in previous sheet

                    # change the sheet name after finished and save into the excel object
                    try:
                        # if the setting name already exist
                        x_try = 0
                        while x_try < c_try:
                            check_sh_temp_name = str(
                                self.new_sheet_name + "_" + str(x_try)
                            )
                            temp_sh = self.excel_ini.wb_res.sheets(check_sh_temp_name)
                            x_try = x_try + 1
                            pass

                        #  this try function must fail and enter except
                        pass

                    except:
                        # if there are no sheet with same name, change the sheet name to related name

                        if x_try == 0:
                            temp_sheet.name = self.new_sheet_name
                        else:
                            temp_sheet.name = str(
                                self.new_sheet_name + "_" + str(x_try)
                            )

                self.excel_ini.ref_table_list[x_sheets] = temp_sheet
                # self.sh_ref_table = self.excel_ini.sh_ref_table

                # # 220914 move to build file to prevent error of delete and end of file
                # # copy the result sheet to result book
                # self.excel_ini.sh_raw_out.copy(self.excel_ini.sh_ref)
                # # assign the sheet to result book
                # self.excel_ini.sh_raw_out = self.excel_ini.wb_res.sheets('raw_out')

                # # copy the sheets to new book
                # # for the new sheet generation, located in sheet_gen
                # self.excel_s.sh_main.copy(self.sh_ref_condition)
                # self.sh_result.copy(self.sh_ref)
                x_sheets = x_sheets + 1
                pass
                # assign to the first sheet after the generation
            self.excel_ini.sh_ref_table = self.excel_ini.ref_table_list[0]
            pass

        pass

    def table_return(self):
        # need to recover this sheet: self.excel_ini.sh_ref_table
        self.excel_ini.sh_ref_table = self.excel_ini.wb.sheets("table")
        # reset the waveform information
        self.excel_ini.wave_condition = ""

        # reset sheet choice to wait for next sheet name update
        self.sheet_name_ready = 0

        pass

    def reload_waveform_dimension(self):
        self.excel_ini.wave_height = self.target_height
        self.excel_ini.wave_width = self.target_width
        print("update the heigh and width settings")

        pass


if __name__ == "__main__":
    #  the testing code for this file object

    # ======== only for object programming
    # testing used temp instrument
    # need to become comment when the OBJ is finished
    import mcu_obj as mcu
    import inst_pkg_d as inst

    # initial the object and set to simulation mode
    pwr_t = inst.LPS_505N(3.7, 0.5, 3, 1, "off")
    pwr_t.sim_inst = 0
    # initial the object and set to simulation mode
    met_v_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_v_t.sim_inst = 0
    load_t = inst.chroma_63600(1, 7, "CCL")
    load_t.sim_inst = 0
    met_i_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_i_t.sim_inst = 0
    src_t = inst.Keth_2440(0, 0, 24, "off", "CURR", 15)
    src_t.sim_inst = 0
    chamber_t = inst.chamber_su242(25, 10, "off", -45, 180, 0)
    chamber_t.sim_inst = 0
    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu_t = mcu.MCU_control(0, 3)

    # for the single test, need to open obj_main first,
    # the real situation is: sheet_ctrl_main_obj will start obj_main first
    # so the file will be open before new excel object benn define

    # and the different verification method can be call below

    version_select = 2

    if version_select == 0:
        # using the main control book as default
        excel_t = par.excel_parameter("obj_main")
        # ======== only for object programming

        # open the result book for saving result
        excel_t.open_result_book()

        # change simulation mode delay (in second)
        excel_t.sim_mode_delay(0.02, 0.01)
        inst.wait_time = 0.01
        inst.wait_samll = 0.01

        # create one file
        format_gen_1 = format_gen(excel_t, "CTRL_sh_ex")

        # generate(or copy) the needed sheet to the result book
        format_gen_1.sheet_gen()

        # start the testing
        format_gen_1.run_format_gen()
        # return the table after data input finished
        format_gen_1.table_return()

        # create one file
        format_gen_2 = format_gen(excel_t, "CTRL_sh_line")

        # generate(or copy) the needed sheet to the result book
        format_gen_2.sheet_gen()

        # start the testing
        format_gen_2.run_format_gen()
        format_gen_2.table_return()

        # create one file
        format_gen_3 = format_gen(excel_t, "CTRL_sh_load")

        # generate(or copy) the needed sheet to the result book
        format_gen_3.sheet_gen()

        # start the testing
        format_gen_3.run_format_gen()
        format_gen_3.table_return()

        # remember that this is only call by main, not by  object
        excel_t.end_of_file(0)

        print("end of the SWIRE scan object testing program")

        pass

    elif version_select == 1:
        # using the main control book as default
        excel_t = par.excel_parameter("obj_main")
        # ======== only for object programming

        # open the result book for saving result
        excel_t.open_result_book()

        # change simulation mode delay (in second)
        excel_t.sim_mode_delay(0.02, 0.01)
        inst.wait_time = 0.01
        inst.wait_samll = 0.01

        format_gen_1 = format_gen(excel_t)
        format_gen_1.set_sheet_name("CTRL_sh_ripple")
        format_gen_1.sheet_gen()
        format_gen_1.run_format_gen()
        # return the table after data input finished
        format_gen_1.table_return()

        format_gen_1.set_sheet_name("CTRL_sh_line")
        format_gen_1.sheet_gen()
        format_gen_1.run_format_gen()
        # return the table after data input finished
        format_gen_1.table_return()

        format_gen_1.set_sheet_name("CTRL_sh_load")
        format_gen_1.sheet_gen()
        format_gen_1.run_format_gen()
        # return the table after data input finished
        format_gen_1.table_return()

        pass

    elif version_select == 2:
        """
        230518 add the format gen used to build big waveform table
        """

        # 230525 sequence of width and heigh seems to be reverse
        width = 145.2
        heigh = 45


        import sheet_ctrl_main_obj as sh

        excel_m = par.excel_parameter(str(sh.file_setting))

        excel_m.open_result_book()
        excel_m.excel_save()

        format_gen_1 = format_gen(excel_m)
        format_gen_1.set_sheet_name("glitch")

        pass
