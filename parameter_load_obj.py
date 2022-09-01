#  this is the object define for parameter loaded

# import for excel control
import xlwings as xw
# this import is for the VBA function
import win32com.client


class excel_parameter ():
    def __init__(self, book_name):
        # when define the object, open the initialization and load all the parameter
        # into the object with related book name
        # this can help to change the setting by giving different book_name to
        # the verification oject
        # each verification object can get different parameter from different
        # excel_parameter object input
        print('start of the parameter loaded')

        self.control_book_trace = 'c:\\py_gary\\test_excel\\' + \
            str(book_name) + '.xlsm'

        # slection for the control parameter mapped to the object
        if book_name != 'obj_main':
            # open the new book for loading the control parameter in related object
            self.wb = xw.Book(self.control_book_trace)
            print('select other book')
            # 220901 first not used the function to switch to other books

            pass
        else:
            # if book loaded is the original control book, not going to open the book and
            # use the obj_main as the control parameter input
            self.wb = xw.books('obj_main.xlsm')
            print('select original book')

            pass

        # after choosing the workbook, define the main sheet to load parameter
        self.sh_main = self.wb.sheets('main')

        # only the instrument control will be still mapped to the original excel
        # since inst_ctrl is no needed to copy to the result sheet
        self.sh_inst_ctrl = self.wb.sheets('inst_ctrl')

        # other way to define sheet:
        # this is the format for the efficiency result
        ex_sheet_name = 'raw_out'
        self.sh_raw_out = self.wb.sheets(ex_sheet_name)
        # this is the sheet for efficiency testing command
        self.sh_volt_curr_cmd = self.wb.sheets('V_I_com')
        # this is the sheet for I2C command
        self.sh_i2c_cmd = self.wb.sheets('I2C_ctrl')

        # file name from the master excel
        self.new_file_name = str(self.sh_main.range('B8').value)
        # load the trace as string
        self.excel_temp = str(self.sh_main.range('B9').value)
        # program control variable (auto and inst settings)
        self.auto_inst_ctrl = self.sh_main.range('B11').value
        # program exit interrupt variable
        self.program_exit = self.sh_main.range('B12').value
        # verification re-run
        self.re_run_verification = self.sh_main.range('B13').value
        # file settings; decide which file to load parameters
        self.file_setting = self.sh_main.range('B14').value
        # program selection setting
        self.program_group_index = self.sh_main.range('B15').value

        # default extra file name is call _temp
        # only multi item case will have temp file
        self.extra_file_name = '_temp'

        # update the result book trace
        self.result_book_trace = self.excel_temp + \
            self.new_file_name + self.extra_file_name + '.xlsx'

        # insturment parameter loading, to load the instrumenet paramenter
        # need to get the index of each item first
        # index need to change if adding new control parameter

        self.index_par_pre_con = self.sh_main.range((3, 9)).value
        self.index_GPIB_inst = self.sh_main.range((4, 9)).value
        self.index_general_other = self.sh_main.range((5, 9)).value
        self.index_pwr_inst = self.sh_main.range((6, 9)).value
        self.index_chroma_inst = self.sh_main.range((3, 12)).value
        self.index_src_inst = self.sh_main.range((4, 12)).value
        self.index_meter_inst = self.sh_main.range((5, 12)).value
        self.index_chamber_inst = self.sh_main.range((6, 12)).value
        self.index_IQ_scan = self.sh_main.range((3, 15)).value
        self.index_eff = self.sh_main.range((4, 15)).value
        # self.index_meter_inst = self.sh_main.range((5, 15)).value
        # self.index_chamber_inst = self.sh_main.range((6, 15)).value

        # base on output format copied from the control book
        # start parameter initialization
        # pre- test condition settings
        self.pre_vin = self.sh_main.range(
            (self.index_par_pre_con + 1, 3)).value
        self.pre_vin_max = self.sh_main.range(
            (self.index_par_pre_con + 2, 3)).value
        self.pre_imax = self.sh_main.range(
            (self.index_par_pre_con + 3, 3)).value
        self.pre_test_en = self.sh_main.range(
            (self.index_par_pre_con + 4, 3)).value
        self.pre_sup_iout = self.sh_main.range(
            (self.index_par_pre_con + 5, 3)).value

        # load the GPIB address for the instrument
        # GPIB instrument list (address loading, name feed back)
        self.pwr_supply_addr = self.sh_main.range(
            (self.index_GPIB_inst + 1, 3)).value
        # met1 usually for voltage
        self.meter1_addr = self.sh_main.range(
            (self.index_GPIB_inst + 2, 3)).value
        # met2 usually for current
        self.meter2_addr = self.sh_main.range(
            (self.index_GPIB_inst + 3, 3)).value
        self.loader_addr = self.sh_main.range(
            (self.index_GPIB_inst + 4, 3)).value
        self.loader_src_addr = self.sh_main.range(
            (self.index_GPIB_inst + 5, 3)).value
        self.temp_chamber_addr = self.sh_main.range(
            (self.index_GPIB_inst + 6, 3)).value

        # initialization for all the object, based on the input parameter of the index

        # parameter setting for the power supply
        self.pwr_vset = self.sh_main.range((self.index_pwr_inst + 1, 3)).value
        self.pwr_iset = self.sh_main.range((self.index_pwr_inst + 2, 3)).value
        self.pwr_act_ch = self.sh_main.range(
            (self.index_pwr_inst + 3, 3)).value
        self.pwr_ini_state = self.sh_main.range(
            (self.index_pwr_inst + 4, 3)).value
        self.relay0_ch = self.sh_main.range((self.index_pwr_inst + 5, 3)).value
        self.relay6_ch = self.sh_main.range((self.index_pwr_inst + 6, 3)).value
        self.relay7_ch = self.sh_main.range((self.index_pwr_inst + 7, 3)).value
        # pre-increase for efficiency measurement
        self.pre_inc_vin = self.sh_main.range(
            (self.index_pwr_inst + 8, 3)).value
        # the setting for Vin calibration accuracy
        self.vin_diff_set = self.sh_main.range(
            (self.index_pwr_inst + 9, 3)).value

        # parameter setting for the chroma loader
        self.loader_act_ch = self.sh_main.range(
            (self.index_chroma_inst + 1, 3)).value
        self.loader_ini_mode = self.sh_main.range(
            (self.index_chroma_inst + 2, 3)).value
        self.loader_cal_ELch = self.sh_main.range(
            (self.index_chroma_inst + 3, 3)).value
        self.loader_cal_VCIch = self.sh_main.range(
            (self.index_chroma_inst + 4, 3)).value
        self.loader_ELch = self.sh_main.range(
            (self.index_chroma_inst + 5, 3)).value
        self.loader_ini_state = self.sh_main.range(
            (self.index_chroma_inst + 6, 3)).value
        self.loader_VCIch = self.sh_main.range(
            (self.index_chroma_inst + 7, 3)).value
        self.loader_cal_mode = self.sh_main.range(
            (self.index_chroma_inst + 8, 3)).value

        # parameter setting for source meter
        self.src_vset = self.sh_main.range((self.index_src_inst + 1, 3)).value
        self.src_iset = self.sh_main.range((self.index_src_inst + 2, 3)).value
        self.src_ini_state = self.sh_main.range(
            (self.index_src_inst + 3, 3)).value
        self.src_ini_type = self.sh_main.range(
            (self.index_src_inst + 4, 3)).value
        self.src_clamp_ini = self.sh_main.range(
            (self.index_src_inst + 5, 3)).value

        # parameter setting for meter
        self.met_v_res = self.sh_main.range(
            (self.index_meter_inst + 1, 3)).value
        self.met_v_max = self.sh_main.range(
            (self.index_meter_inst + 2, 3)).value
        self.met_i_res = self.sh_main.range(
            (self.index_meter_inst + 3, 3)).value
        self.met_i_max = self.sh_main.range(
            (self.index_meter_inst + 4, 3)).value

        # parameter setting for chamber

        self.cham_tset_ini = self.sh_main.range(
            (self.index_chamber_inst + 1, 3)).value
        self.cham_ini_state = self.sh_main.range(
            (self.index_chamber_inst + 2, 3)).value
        self.cham_l_limt = self.sh_main.range(
            (self.index_chamber_inst + 3, 3)).value
        self.cham_h_limt = self.sh_main.range(
            (self.index_chamber_inst + 4, 3)).value
        self.cham_hyst = self.sh_main.range(
            (self.index_chamber_inst + 5, 3)).value

        # other control parameter
        # COM port parameter input
        self.mcu_com_addr = self.sh_main.range(
            self.index_general_other + 1, 3).value
        # general delay time
        self.wait_time = self.sh_main.range(
            self.index_general_other + 2, 3).value
        # the start point for the raw_out index
        self.raw_y_position_start = self.sh_main.range(
            self.index_general_other + 3, 3).value
        self.raw_x_position_start = self.sh_main.range(
            self.index_general_other + 4, 3).value
        self.sheet_off_finished = self.sh_main.range(
            self.index_general_other + 5, 3).value
        # plot pause control (1 is enable, 0 is disable)
        self.en_plot_waring = self.sh_main.range(
            self.index_general_other + 6, 3).value
        self.en_fully_auto = self.sh_main.range(
            self.index_general_other + 7, 3).value
        self.en_start_up_check = self.sh_main.range(
            self.index_general_other + 8, 3).value

        # verification item: IQ parameter
        self.ISD_range = self.sh_main.range(
            self.index_IQ_scan + 1, 3).value

        # verification item: eff control parameter
        self.channel_mode = self.sh_main.range(self.index_eff + 1, 3).value
        # SWIRE or I2C selected setting
        self.sw_i2c_select = self.sh_main.range(self.index_eff + 2, 3).value
        # if the channel 1=> EL power, 2=> AVDD, 0=> not use source meter
        # when control = 0, all channel used chroma's output mapping
        self.source_meter_channel = self.sh_main.range(
            self.index_eff + 3, 3).value

        print('end of the parameter loaded')

        pass

    def open_result_book(self):

        # define new result workbook
        self.wb_res = xw.Book()
        # create reference sheet (for sheet position)
        # sh_ref is the index for result sheet
        # sh_ref_condition is for testing condition and setting
        # all the reference sheet will delete after the program finished
        self.sh_ref = self.wb_res.sheets.add('ref_sh')
        self.sh_ref_condition = self.wb_res.sheets.add('ref_sh2')
        # delete the extra sheet from new workbook, difference from version
        self.wb_res.sheets('工作表1').delete()

        # copy the sheets to new book
        self.sh_main.copy(self.sh_ref)
        # assign both sheet to the new sheets in result book
        self.sh_main = self.wb_res.sheets('main')

        # for the other sheet rather than main, will decide to copy to result
        # or not depends on verification item is used or not
        pass

    def end_of_test(self):
        # at the end of test, delete the reference sheet and save the file

        self.sh_ref.delete()
        self.sh_ref_condition.delete()
        # update the result book trace
        # extra file name should be update by the last item or the single item
        self.result_book_trace = self.excel_temp + \
            self.new_file_name + self.extra_file_name + '.xlsx'
        self.wb_res.save(self.result_book_trace)
        pass

    def save(self):
        # save with the name end with temp
        self.wb_res.save(self.result_book_trace)
        pass

    pass


if __name__ == '__main__':
    #  the testing code for this file object
    excel = excel_parameter('obj_main')

    input()

    excel2 = excel_parameter('other_testing_condition')

    input()
