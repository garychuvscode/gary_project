#  this is the object define for parameter loaded

# import for excel control
import xlwings as xw
# this import is for the VBA function
import win32com.client
# application of array
import numpy as np
# include for atof function => transfer string to float
import locale as lo


# ======== excel application related
# 開啟 Excel 的app
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
# ======== excel application related


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

        # 220907 add the sheet array for eff measurement
        self.sheet_arry = np.full([200], None)

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
        # this is the sheet for IQ scan
        self.sh_iq_scan = self.wb.sheets('IQ_measured')
        # this is the sheet for wire scan
        self.sh_sw_scan = self.wb.sheets('SWIRE_scan')

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

        # Vin status global variable
        self.vin_status = ''
        # I_AVDD status global variable
        self.i_avdd_status = ''
        # I_EL staatus global variable
        self.i_el_status = ''
        # SW_I2C status global variable
        self.sw_i2c_status = ''

        # default extra file name is call _temp
        # only multi item case will have temp file
        self.extra_file_name = '_temp'
        # for report contain one or multi testing items
        # default is single ( one verification)
        # change the name to _pXX for multi program
        # XX means the program selection number

        # 220907 add another variable call detail name for the
        # eff or I2C measurement
        # since save from each round of the eff test file
        self.detail_name = ''

        # result_book_trace change in the sub_program
        # update the result book trace
        self.full_result_name = self.new_file_name + \
            self.extra_file_name + self.detail_name
        self.result_book_trace = self.excel_temp + \
            self.new_file_name + self.extra_file_name + self.detail_name + '.xlsx'

        # insturment parameter loading, to load the instrumenet paramenter
        # need to get the index of each item first
        # index need to change if adding new control parameter

        self.index_par_pre_con = 0
        self.index_GPIB_inst = 0
        self.index_general_other = 0
        self.index_pwr_inst = 0
        self.index_chroma_inst = 0
        self.index_src_inst = 0
        self.index_meter_inst = 0
        self.index_chamber_inst = 0
        self.index_IQ_scan = 0
        self.index_eff = 0

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

        # index check put at the open result sheet

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
        # met1v usually for voltage
        self.meter1_v_addr = self.sh_main.range(
            (self.index_GPIB_inst + 2, 3)).value
        # met2 usually for current
        self.meter2_i_addr = self.sh_main.range(
            (self.index_GPIB_inst + 3, 3)).value
        self.loader_addr = self.sh_main.range(
            (self.index_GPIB_inst + 4, 3)).value
        self.loader_src_addr = self.sh_main.range(
            (self.index_GPIB_inst + 5, 3)).value
        self.chamber_addr = self.sh_main.range(
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
        self.book_off_finished = self.sh_main.range(
            self.index_general_other + 5, 3).value
        # plot pause control (1 is enable, 0 is disable)
        self.en_plot_waring = self.sh_main.range(
            self.index_general_other + 6, 3).value
        self.en_fully_auto = self.sh_main.range(
            self.index_general_other + 7, 3).value
        self.en_start_up_check = self.sh_main.range(
            self.index_general_other + 8, 3).value
        self.wait_small = self.sh_main.range(
            self.index_general_other + 9, 3).value

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
        # when control = 0, all channel used chroma's output mapping
        self.eff_chamber_en = self.sh_main.range(
            self.index_eff + 4, 3).value
        self.eff_single_file = self.sh_main.range(
            self.index_eff + 5, 3).value

        # add the loop control for each items

        # counteer is usually use c_ in opening

        # EFF_inst used
        self.c_avdd_load = self.sh_volt_curr_cmd.range('D1').value
        self.c_vin = self.sh_volt_curr_cmd.range('B1').value
        self.c_iload = self.sh_volt_curr_cmd.range('C1').value
        self.c_pulse = self.sh_volt_curr_cmd.range('E1').value
        self.c_i2c = self.sh_i2c_cmd.range('B1').value
        self.c_i2c_g = self.sh_i2c_cmd.range('D1').value
        self.c_avdd_single = self.sh_volt_curr_cmd.range('G1').value
        self.c_avdd_pulse = self.sh_volt_curr_cmd.range('H1').value
        self.c_tempature = self.sh_volt_curr_cmd.range('I1').value

        # IQ testing related
        self.c_iq = self.sh_iq_scan.range('C4').value
        self.iq_scaling = self.sh_iq_scan.range('C5').value

        # SWIRE_scan  related
        self.c_swire = self.sh_sw_scan.range('B2').value
        self.vin_set = self.sh_sw_scan.range('C6').value
        self.Iin_set = self.sh_sw_scan.range('E6').value
        self.EL_curr = self.sh_sw_scan.range('C7').value
        self.VCI_curr = self.sh_sw_scan.range('E7').value

        # efficiency test needed variable
        self.eff_done_sh = 0
        self.sub_sh_count = 0
        self.one_file_sheet_adj = 0

        # temp sheet name for the plot index
        # because there are different mode, no need specific channel name, just the positive and negative
        self.eff_temp = ''
        self.pos_temp = ''
        self.neg_temp = ''
        self.raw_temp = ''
        # 220825 add for vout and von regulation
        self.pos_pre_temp = ''
        self.neg_pre_temp = ''

        # the excel table gap for the data in raw sheet
        self.raw_gap = 4 + 10
        # setting of raw gap is the " gap + element "
        # single eff: Vin, Iin, Vout, Iout, Eff => 5 elements

        # active sheet (for the result and raw)
        self.sheet_active = ''
        # for efficiency
        self.raw_active = ''
        # for raw data
        self.vout_p_active = ''
        # for ELVDD, or AVDD
        self.vout_n_active = ''
        # for ELVSS
        self.vout_p_pre_active = ''
        # for VOP
        self.vout_n_pre_active = ''
        # for VON

        print('end of the parameter loaded')

        pass

    def open_result_book(self):

        # before open thr result book to check index, first check and correct
        # index (will be update to the obj_main)
        self.index_check()

        # define new result workbook
        self.wb_res = xw.Book()
        # create reference sheet (for sheet position)
        # sh_ref is the index for result sheet
        # sh_ref_condition is for testing condition and setting
        # all the reference sheet will delete after the program finished
        self.sh_ref = self.wb_res.sheets.add('ref_sh')
        # self.sh_ref_condition = self.wb_res.sheets.add('ref_sh2')
        # delete the extra sheet from new workbook, difference from version
        self.wb_res.sheets('工作表1').delete()

        # copy the main sheets to new book
        self.sh_main.copy(self.sh_ref)
        # assign sheet to the new sheets in result book
        self.sh_main = self.wb_res.sheets('main')

        # for the other sheet rather than main, will decide to copy to result
        # or not depends on verification item is used or not
        pass

    def end_of_file(self, multi_items):
        # at the end of test, delete the reference sheet and save the file
        # 220907 change name from end_of_test to end_of file, since the operation
        # to cut file may be needed during single test

        self.sh_ref.delete()
        # self.sh_ref_condition.delete()
        # update the result book trace
        # extra file name should be update by the last item or the single item
        if multi_items == 1:
            # using multi item extra file name
            self.extra_file_name = '_p' + str(int(self.program_group_index))



        self.result_book_trace = self.excel_temp + \
            self.new_file_name + self.extra_file_name + self.detail_name + '.xlsx'
        self.full_result_name = self.new_file_name + \
            self.extra_file_name + self.detail_name
        self.wb_res.save(self.result_book_trace)

        if self.book_off_finished == 1:
            self.wb_res.close()
            pass

        # to reset the sheet after file finished and turn off
        self.sheet_reset()
        self.detail_name = ''
        self.extra_file_name = '_temp'
        self.new_file_name = str(self.sh_main.range('B8').value)
        self.full_result_name = self.new_file_name + \
            self.extra_file_name + self.detail_name
        self.result_book_trace = self.excel_temp + \
            self.new_file_name + self.extra_file_name + self.detail_name + '.xlsx'

        # reset the sheet count of the one file efficiency when end of file
        self.one_file_sheet_adj = 0
        self.sh_temp.delete()

        pass

    def excel_save(self):
        # only save, not change the result book trace
        # should be only the temp file during program operation
        self.wb_res.save(self.result_book_trace)
        pass

    def inst_name_sheet(self, nick_name, full_name):
        # definition of sub program may not need the self, but definition of class will need the self
        # self is usually used for internal parameter of class
        # this function will get the nick name and full name from main and update to the sheet
        # based on the nick name

        # 220902 operate after main is change, and result will be correct

        if nick_name == 'PWR1':
            self.sh_main.range(
                (self.index_GPIB_inst + 1, 4)).value = full_name

        elif nick_name == 'MET1':
            self.sh_main.range(
                (self.index_GPIB_inst + 2, 4)).value = full_name

        elif nick_name == 'MET2':
            self.sh_main.range(
                (self.index_GPIB_inst + 3, 4)).value = full_name

        elif nick_name == 'LOAD1':
            self.sh_main.range(
                (self.index_GPIB_inst + 4, 4)).value = full_name

        elif nick_name == 'LOADSR':
            self.sh_main.range(
                (self.index_GPIB_inst + 5, 4)).value = full_name

        elif nick_name == 'chamber':
            self.sh_main.range(
                (self.index_GPIB_inst + 6, 4)).value = full_name

        pass

    def sheet_reset(self):
        # this sheet reset the all the sheet variable assignment to the original sheet
        # in main_obj, which used for the re-run program
        # just copoy from the sheet assignment

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
        # this is the sheet for IQ scan
        self.sh_iq_scan = self.wb.sheets('IQ_measured')

        pass

    def index_check(self):
        # this sub program used to check the index setting for the excel input
        # prevent logic error of the wrong indexing of program parameter
        check_str = 'settings'
        index_correction = 0
        item_array = np.full([10], 1)
        check_ctrl = self.sum_array(item_array)

        while(check_ctrl > 0):
            # while there are items not pass, need to re-run check parameter
            # porcess

            if self.sh_main.range((int(self.index_par_pre_con), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_par_pre_con check done')
                item_array[0] = 0
                print(item_array)
                pass
            else:
                print('index_par_pre_con check fail')
                print('correct value in: (3, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((3, 9)).value = index_correction
                self.index_par_pre_con = index_correction
                pass

            if self.sh_main.range((int(self.index_GPIB_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_GPIB_inst check done')
                item_array[1] = 0
                print(item_array)
                pass
            else:
                print('index_GPIB_inst check fail')
                print('correct value in: (4, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((4, 9)).value = index_correction
                self.index_GPIB_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_general_other), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_general_other check done')
                item_array[2] = 0
                print(item_array)
                pass
            else:
                print('index_general_other check fail')
                print('correct value in: (5, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((5, 9)).value = index_correction
                self.index_general_other = index_correction
                pass

            if self.sh_main.range((int(self.index_pwr_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_pwr_inst check done')
                item_array[3] = 0
                print(item_array)
                pass
            else:
                print('index_pwr_inst check fail')
                print('correct value in: (6, 9)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((6, 9)).value = index_correction
                self.index_pwr_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_chroma_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_chroma_inst check done')
                item_array[4] = 0
                print(item_array)
                pass
            else:
                print('index_chroma_inst check fail')
                print('correct value in: (3, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((3, 12)).value = index_correction
                self.index_chroma_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_src_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_src_inst check done')
                item_array[5] = 0
                print(item_array)
                pass
            else:
                print('index_src_inst check fail')
                print('correct value in: (4, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((4, 12)).value = index_correction
                self.index_src_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_meter_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_meter_inst check done')
                item_array[6] = 0
                print(item_array)
                pass
            else:
                print('index_meter_inst check fail')
                print('correct value in: (5, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((5, 12)).value = index_correction
                self.index_meter_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_chamber_inst), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_chamber_inst check done')
                item_array[7] = 0
                print(item_array)
                pass
            else:
                print('index_chamber_inst check fail')
                print('correct value in: (6, 12)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((6, 12)).value = index_correction
                self.index_chamber_inst = index_correction
                pass

            if self.sh_main.range((int(self.index_IQ_scan), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_IQ_scan check done')
                item_array[8] = 0
                print(item_array)
                pass
            else:
                print('index_IQ_scan check fail')
                print('correct value in: (3, 15)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((3, 15)).value = index_correction
                self.index_IQ_scan = index_correction
                pass

            if self.sh_main.range((int(self.index_eff), 3)).value == check_str:
                # index pass if value is loaded as settings
                print('index_eff check done')
                item_array[9] = 0
                print(item_array)
                pass
            else:
                print('index_eff check fail')
                print('correct value in: (4, 15)')
                print('the new index number input:')
                index_correction = lo.atof(input())
                self.sh_main.range((4, 15)).value = index_correction
                self.index_eff = index_correction
                pass
            # update while loop value
            check_ctrl = self.sum_array(item_array)

            pass

        print('the index correction finished!')
        # save the wb for the new index correction settings
        # the wb in excel object is mapped to the obj_main file
        self.wb.save()
        pass

    def sum_array(self, arr):
        # initialize a variable
        # to store the sum
        # while iterating through
        # the array later
        sum = 0

        # iterate through the array
        # and add each element to the sum variable
        # one at a time
        for i in arr:
            sum = sum + i

        return(sum)
        pass

    def sim_mode_delay(self, wait_time, wait_small):
        # reduce the delay time for the simulation mode
        self.wait_time = wait_time
        self.wait_time = wait_small
        pass

    # SWIRE request sub-program

    def ideal_v_table(self, c_swire):
        ideal_v_res = self.sh_sw_scan.range((11 + c_swire, 3)).value
        return ideal_v_res

    #  efficiency testing sub-program

    def build_file(self, detail_name):

        # 220907 mapped variable with excel object
        self.detail_name = detail_name
        wb = self.wb
        wb_res = self.wb_res
        result_sheet_name = 'raw_out'
        new_file_name = self.new_file_name
        excel_trace = self.excel_temp
        channel_mode = self.channel_mode
        c_avdd_load = self.c_avdd_load
        # this sheet mapped to raw out at eff_inst
        sh_org_tab2 = self.sh_raw_out
        # this sheet mapped to volage and current command
        sh_org_tab = self.sh_volt_curr_cmd
        sh_ref = self.sh_ref
        sheet_arry = self.sheet_arry

        # cpoy the sheet to the result book, will be set at the eff_obj
        # 220907: here is only for the book generation, must call the sheet
        # generation before call the build file
        # assign for the result sheet to the excel object will also be done in the
        # sheet gen of eff_obj

        # update the file name, but not update file name untill the file is finished

        # save the result book and turn off the control book
        wb_res.save(self.result_book_trace)
        # wb.close()
        # close control books

        # base on output format copied from the control book
        # start parameter initialization

        # used array to setup the result sheet
        # res_sheet_array = np.zeros(100)
        # this method not working, skip this time

        # here is to generate the sheet
        # counter selection for sheet generation: EL power only one cycle needed (IAVDD= 0)
        if channel_mode == 0 or channel_mode == 1:
            # both only EL or only AVDD just one time
            c_sheet_copy = 1
            if channel_mode == 1:
                sub_sh_count = 4

                # eff + raw + AVDD regulation + Vout regulation
                # 220825 add vout sheet
            elif channel_mode == 0:
                sub_sh_count = 6
                # eff + raw + ELVDD regulation + ELVSS regulation + Vout regulation + Von regulation
        elif channel_mode == 2:
            c_sheet_copy = c_avdd_load
            sub_sh_count = 6
            # eff + raw + ELVDD regulation + ELVSS regulation + Vout regulation + Von regulation

        self.sub_sh_count = sub_sh_count
        # 220911 update the sub sh count after confirm the parameter

        x_sheet_copy = 0
        sh_temp = sh_org_tab2
        self.sh_temp = sh_temp
        # this loop build the extra sheet needed in the program
        # there are one raw data sheet and fixed format summarize table
        # need to build both efficiency and load regulation summarize table in this loop
        # 3-channel efficiency
        while x_sheet_copy < c_sheet_copy:
            # issue: if x_sheet_copy == 0:

            if channel_mode == 2 or channel_mode == 0:
                # sheet needed for 3 chand only EL are the same
                # just define the different sheet name

                # load AVDD current parameter
                excel_temp = str(sh_org_tab.range(3 + x_sheet_copy, 4).value)

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'EFF_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'EFF'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                if channel_mode == 2:
                    # add another sheet for the raw data of each AVDD current
                    sheet_temp = 'RAW_I_AVDD=' + excel_temp + 'A'
                else:
                    sheet_temp = 'RAW'
                # raw data sheet no need example format, can use empty sheet
                sheet_arry[sub_sh_count * x_sheet_copy + 1] = sheet_temp
                wb_res.sheets.add(sheet_temp)
                # =======
                # 220825 explanation added: since the sheet of raw data doesn't have specific
                # format and input needed, add the sheet directly, no need to copy
                # this is the reason why it's different with other sheet generation
                # to add the Vout and Von load regulation, use the format in excel raw_out
                # and it's general format for the regulation and plot function in VBA

                # =======
                # add another sheet for the ELVDD data of each AVDD current
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'ELVDD_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'ELVDD'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheset name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 2] = sheet_temp
                sh_org_tab2.name = sheet_temp
                # =======

                # =======
                # add another sheet for the ELVSS data of each AVDD current
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'ELVSS_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'ELVSS'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 3] = sheet_temp
                sh_org_tab2.name = sheet_temp
                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'Vop_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'Vop'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 4] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving
                if channel_mode == 2:
                    # 3-ch operation
                    sheet_temp = 'Von_I_AVDD=' + excel_temp + 'A'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                        3 + x_sheet_copy, 4).value
                else:
                    # EL operation
                    sheet_temp = 'Von'
                    # assign the AVDD settting to blue blank of the sheet
                    sh_org_tab2.range(21, 3).value = '0'
                    # no AVDD current, but channel turn on in this operation
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 5] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

            elif channel_mode == 1:
                # sheet build up for only AVDD

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving

                # AVDD operation
                sheet_temp = 'EFF'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = 'NA'
                # here is for AVDD eff
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                # add another sheet for the raw data of each AVDD current
                sheet_temp = 'RAW'
                # raw data sheet no need example format, can use empty sheet
                sheet_arry[sub_sh_count * x_sheet_copy + 1] = sheet_temp
                wb_res.sheets.add(sheet_temp)
                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving

                # AVDD operation
                sheet_temp = 'AVDD'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = 'NA'
                # here is for AVDD eff
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 2] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======

                # =======
                sh_temp.copy(sh_ref)
                sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
                # here is to open a new sheet for data saving

                # AVDD operation
                sheet_temp = 'Vout'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = 'NA'
                # here is for AVDD eff
                # save the sheet name into the array for loading
                sheet_arry[sub_sh_count * x_sheet_copy + 3] = sheet_temp
                sh_org_tab2.name = sheet_temp

                # =======
                self.sub_sh_count = sub_sh_count

            x_sheet_copy = x_sheet_copy + 1

        # sh_temp.delete()
        # don't need the original raw output format, remove the output

    def eff_rerun(self):
        self.eff_done_sh
        self.sh_volt_curr_cmd
        self.sh_raw_out
        self.sh_i2c_cmd
        self.sh_inst_ctrl
        # this program check the status of the excel file eff_re-run block
        # and update the eff_done to restart efficienct testing
        # from the main, this sub will run if eff_done is already 1
        eff_reset_temp = self.sh_main.range('B13').value
        print('wait for re-run, update command and setup then set re-run to 1')
        print('the program will start again')

        if eff_reset_temp == 1:
            eff_done_sh = 0
            # reset to 0 if eff sheet is ready to re-run
            # also need to set te input blank back to 0
            self.sh_main.range('B13').value = 0
            # other wise there will be infinite loop

            # also need to re-assign the mapping sheet to Eff_inst
            # the sheet assignment is gone after finished one round
            self.re_assign_sheet()

            pass
        else:
            # no need for the action of changing the reset status
            pass

        pass

    def check_inst_update(self):

        pass

    def program_status(self, status_string):
        # transfer to the string for following operation
        # if you need to modify in sub-program, need to use global definition
        # global vin_status
        # global i_avdd_status
        # global i_el_status
        # global sw_i2c_status
        status_sting_sub = str(status_string)
        self.sh_main.range((3, 2)).value = status_sting_sub
        # Vin status
        self.sh_main.range('F3').value = self.vin_status
        # I_AVDD status
        self.sh_main.range('F4').value = self.i_avdd_status
        # I_EL staatus
        self.sh_main.range('F5').value = self.i_el_status
        # SW_I2C status
        self.sh_main.range('F6').value = self.sw_i2c_status

        print('status_update: ' + status_sting_sub)
        print(str(self.vin_status) + '-' + str(self.i_avdd_status) +
              '-' + str(self.i_el_status) + '-' + str(self.sw_i2c_status))
        # use for debugging for the program status update
        # input()
        pass

    def act_sheet_loaded(self):
        #  to consider to put the update of sheet assign to here or stay in eff_obj
        pass

    def eff_calculated(self):
        self.value_eff = ((self.value_elvdd - self.value_elvss) * self.value_iel +
                          self.value_avdd * self.value_iavdd) / (self.value_vin * self.value_iin)
        return self.value_eff

    def sheet_adj_for_eff(self, x_avdd):
        # used to adjust the sheet name to prevent conflict when
        # building in single file

        # re-name the sheet with new name

        x_sub_sh_count = 0
        while x_sub_sh_count < self.sub_sh_count:
            index = self.sub_sh_count * x_avdd + x_sub_sh_count
            target_sheet = self.wb_res.sheets(self.sheet_arry[index])
            new_sheet_name = str(self.one_file_sheet_adj) + \
                '_' + self.sheet_arry[index]
            target_sheet.name = new_sheet_name
            self.sheet_arry[index] = new_sheet_name

            # also need to update operating condition to each sheet
            self.condition_note = self.extra_file_name + self.detail_name
            if x_sub_sh_count == 0:
                target_sheet.range('M2').value = 'operating condition'
            target_sheet.range('M3').value = self.condition_note

            x_sub_sh_count = x_sub_sh_count + 1

        self.one_file_sheet_adj = self.one_file_sheet_adj + 1
        pass

    # used for the loading the data to related excel sheet and blank

    def data_latch(self, data_name, mea_res, x_vin, x_iload, value_i_offset1, value_i_offset2):
        raw_gap = self.raw_gap
        channel_mode = self.channel_mode
        # define the globa variable for eff calculation
        # global value_elvdd
        # global value_elvss
        # global value_avdd
        # global value_iin
        # global value_vin
        # global value_iel
        # global value_iavdd

        # global bypass_measurement_flag
        # first to check if the bypass flag raise ~
        # set measurement result to 0 if the bpass flag is enable
        # if bypass_measurement_flag == 1:
        #     mea_res = '0'

        if data_name == 'vin':
            # vin only record in the raw data
            self.raw_active.range((11 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.value_vin = float(mea_res)

        elif data_name == 'iin':
            # iin only record in the raw data
            self.raw_active.range((12 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.value_iin = float(mea_res)

        elif data_name == 'elvdd':
            # elvdd record in the raw data, elvdd regulation
            self.raw_active.range((13 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            # sheet_active.range((25 + x_iload, 3 + x_vin)).value = lo.atof(mea_res)
            if channel_mode == 0 or channel_mode == 2:
                self.vout_p_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)
            self.value_elvdd = float(mea_res)

        elif data_name == 'elvss':
            # elvss record in the raw data, elvss regulation
            self.raw_active.range((14 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            if channel_mode == 0 or channel_mode == 2:
                self.vout_n_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)
            self.value_elvss = float(mea_res)

        elif data_name == 'i_el':
            # i_el only record in the raw data
            self.raw_active.range((15 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res) - value_i_offset1
            self.value_iel = float(mea_res) - value_i_offset1

        elif data_name == 'avdd':
            # avvdd record in the raw data, avvdd regulation
            self.raw_active.range((16 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.value_avdd = float(mea_res)
            # the raw data of AVDD need to record no matter in AVDD only mode or the 3-ch mode
            # the selection of EL only or not is decidde inthe main program
            # here is only for the choice of AVDD regulation
            if channel_mode == 1:
                # only need to record the regulation when is operating for AVDD only mode
                self.vout_p_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)

        elif data_name == 'i_avdd':
            # i_avdd only record in the raw data
            self.raw_active.range((17 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res) - value_i_offset2
            self.value_iavdd = float(mea_res) - value_i_offset2

        elif data_name == 'eff':
            # eff record in the raw data
            self.raw_active.range((18 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.sheet_active.range((25 + x_iload, 3 + x_vin)
                                    ).value = lo.atof(mea_res)

        elif data_name == 'vop':
            # vop record in the raw data
            self.raw_active.range((19 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.vout_p_pre_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)

        elif data_name == 'von':
            # von record in the raw data
            self.raw_active.range((20 + raw_gap * x_vin, 3 + x_iload)
                                  ).value = lo.atof(mea_res)
            self.vout_n_pre_active.range((25 + x_iload, 3 + x_vin)
                                         ).value = lo.atof(mea_res)

        # clear the bypass flag every time enter data latch function
        # bypass_measurement_flag = 0

    # the instrument update check for excel
    def check_inst_update(self):

        pass

    # to plot the result in different sheet
    def plot_single_sheet(self, v_cnt, i_cnt, sheet_n):

        book_n = str(self.full_result_name) + '.xlsx'
        # plot the sheet based on the input sheet name and element length
        excel.Application.Run("obj_main.xlsm!gary_chart",
                              v_cnt, i_cnt, sheet_n, book_n, self.raw_y_position_start, self.raw_x_position_start)
        print('the plot of ' + str(sheet_n) +
              ' in book ' + str(book_n) + ' is finished')
        pass


if __name__ == '__main__':
    #  the testing code for this file object
    excel = excel_parameter('obj_main')

    input()

    excel2 = excel_parameter('other_testing_condition')

    input()
