# this file will create for the general testing
# instrument control is set from main
# 221208: this file also include the other testing related with general test command table
# the plan to add: 1. power on and off(high low temp) 2. waveform capture(inrush current, power on off sequence)


# excel parameter and settings
import parameter_load_obj as par
# for the jump out window
# # also for the jump out window, same group with win32con
import win32api
from win32con import MB_SYSTEMMODAL
# for the delay function
import time
# include for atof function => transfer string to float
import locale as lo


class general_test ():

    def __init__(self, excel0, pwr0, met_v0, loader_0, mcu0, src0, met_i0, chamber0):
        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst
            # initial the object and set to simulation mode
            pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
            pwr0.sim_inst = 0
            # initial the object and set to simulation mode
            met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
            met_v0.sim_inst = 0
            loader_0 = inst.chroma_63600(1, 7, 'CCL')
            loader_0.sim_inst = 0
            # mcu is also config as simulation mode
            mcu0 = mcu.MCU_control(0, 3)
            # using the main control book as default
            excel0 = par.excel_parameter('obj_main')
            src0 = inst.Keth_2440(0, 0, 24, 'off', 'CURR', 15)
            src0.sim_inst = 0
            met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
            met_i0.sim_inst = 0
            chamber0 = inst.chamber_su242(25, 10, 'off', -45, 180, 0)
            chamber0.sim_inst = 0
            # ======== only for object programming

        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        # assign the input information to object variable
        self.excel_ini = excel0
        self.pwr_ini = pwr0
        self.loader_ini = loader_0
        self.met_v_ini = met_v0
        self.mcu_ini = mcu0
        self.src_ini = src0
        self.met_i_ini = met_i0
        self.chamber_ini = chamber0
        # self.single_ini = single0

        # # setup extra file name if single verification
        # if self.single_ini == 0:
        #     # this is not single item verififcation
        #     # and this is not the last item (last item)
        #     pass
        # elif self.single_ini == 1:
        #     # it's single, using it' own file name
        #     # item can decide the extra file name is it's the only item
        #     self.excel_ini.extra_file_name = '_SWIRE_pulse'
        #     pass

        self.excel_ini.extra_file_name = '_general'
        # sheet ready or not index
        self.sheet_name_ready = 0
        # set to 0 for simulation mode
        self.obj_sim_mode = 1

        # 221215: add the self sheet assignment for more extra sheet and record
        # in reservation, use the reference sheet first

        self.ex_sh_array = [0, 0, 0, 0, 0]
        self.extra_count = 0
        # extra count is used to check for the available sequence of extra sheet

        # 221226: add the control index record for the general tesitng
        # default will all be 0
        self.ctrl_ind_1 = 0
        self.ctrl_ind_2 = 0

        # 221226: add the supply iout setting in different channel
        self.iout_r0 = 0.1
        self.iout_r6 = 0.1
        self.iout_r7 = 0.1

        # iout_read for the pwr
        self.pwr_iout_en = 0

    pass

    def extra_file_name_setup(self, extra_name=0):
        if extra_name == 0:
            self.excel_ini.extra_file_name = '_general'
        else:
            self.excel_ini.extra_file_name = '_general' + str(extra_name)
        pass

    def run_verification(self, vin_cal=1, ctrl_ind_1=0, pwr_iout=0, ctrl_ind_2=0):
        '''
        run the general testing: default calibrate Vin on\n
        to disable, change vin_cal to 0\n
        ctrl_ind_1 enable the function before measurement or not (run_veri_add_in_1)\n
        pwr_iout is to record pwr iout or not
        ctrl_ind_2 enable the function after measurement or not (run_veri_add_in_2)\n
        '''
        # 221201 sheet generation move to set_sheet_name
        # # give the sheet generation
        # self.sheet_gen()

        # slave object in subprogram
        pwr_s = self.pwr_ini
        load_s = self.loader_ini
        met_v_s = self.met_v_ini
        mcu_s = self.mcu_ini
        excel_s = self.excel_ini
        load_src_s = self.src_ini
        met_i_s = self.met_i_ini
        chamber_s = self.chamber_ini

        # things must have in all run_verification
        pre_vin = excel_s.pre_vin
        pre_sup_iout = excel_s.pre_sup_iout
        pre_imax = excel_s.pre_imax
        pre_vin_max = excel_s.pre_vin_max

        en_start_up_check = excel_s.en_start_up_check
        pre_test_en = excel_s.pre_test_en
        relay0_ch = excel_s.relay0_ch

        chamber_default_tset = excel_s.cham_tset_ini

        # general testing parameter
        gen_chamber_mea = excel_s.gen_chamber_en
        gen_loader_en = excel_s.gen_loader_en
        gen_met_i_en = excel_s.gen_met_i_en
        gen_volt_ch_amount = excel_s.gen_volt_ch_amount
        gen_pulse_i2x_en = excel_s.gen_pulse_i2x_en
        gen_loader_ch_amount = excel_s.gen_loader_ch_amount
        gen_pwr_ch_amount = excel_s.gen_pwr_ch_amount
        gen_pwr_i_set = excel_s.gen_pwr_i_set
        gen_col_amount = excel_s.gen_col_amount
        # 221111 disable extra enable control, use the simulation mode of GPIB to
        # control the instrument

        # make sure MCU back to initial
        self.mcu_ini.back_to_initial()

        # 221220: add the function for reading pwr supply iout
        # 0 is disable and 1 is enable => just check for data latch
        # add the measurement for all case
        self.pwr_iout_en = pwr_iout

        # power supply OV and OC protection
        pwr_s.ov_oc_set(pre_vin_max, pre_imax)

        # power supply channel (channel on setting)
        if pre_test_en == 1:
            if gen_pwr_ch_amount >= 1:
                pwr_s.chg_out(
                    pre_vin, pre_sup_iout, excel_s.relay0_ch, 'on')
            if gen_pwr_ch_amount > 1:
                pwr_s.chg_out(
                    pre_vin, pre_sup_iout, excel_s.relay6_ch, 'on')
            if gen_pwr_ch_amount > 2:
                pwr_s.chg_out(
                    pre_vin, pre_sup_iout, excel_s.relay7_ch, 'on')
            print('pre-power on here')
            # turn off the power and load

            print('also turn all load off')

            if en_start_up_check == 1:
                excel_s.message_box('press enter if hardware configuration is correct',
                                    'Pre-power on for system test under Vin= ' + str(pre_vin) + 'Iin= ' + str(pre_sup_iout))
                # msg_res = win32api.MessageBox(
                #     0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(pre_vin) + 'Iin= ' + str(pre_sup_iout))

            if gen_chamber_mea == 1:
                # chamber turn on with default setting, using default temperature
                # chamber_s.chamber_set(chamber_default_tset)
                print('chamber enable')

        x_count = 0
        while x_count < self.c_test_amount:
            # 221220: this control index used for other function in object
            self.x_count = x_count

            # load the setting first
            self.data_loaded(x_count)

            # if gen_chamber_mea == 1:
            # set up chamber temperature
            if self.chamber_target != 'x':
                self.res_temp_read = chamber_s.chamber_set(self.chamber_target)

            # pwr_s.change_V(self.pwr_ch1, 1)
            # pwr_s.change_V(self.pwr_ch2, 2)
            # pwr_s.change_V(self.pwr_ch3, 3)
            # if gen_pwr_ch_amount >= 1:
            if gen_pwr_ch_amount >= 1:
                if self.pwr_ch1 != 'x':
                    pwr_s.chg_out(
                        self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    pwr_s.chg_out(0, self.iout_r0,
                                  self.excel_ini.relay0_ch, 'off')
            if gen_pwr_ch_amount > 1:
                if self.pwr_ch2 != 'x':
                    pwr_s.chg_out(
                        self.pwr_ch2, self.iout_r6, self.excel_ini.relay6_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    pwr_s.chg_out(0, self.iout_r6,
                                  self.excel_ini.relay6_ch, 'off')
            if gen_pwr_ch_amount > 2:
                if self.pwr_ch3 != 'x':
                    pwr_s.chg_out(
                        self.pwr_ch3, self.iout_r7, self.excel_ini.relay7_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    pwr_s.chg_out(0, self.iout_r7,
                                  self.excel_ini.relay7_ch, 'off')

            if gen_pulse_i2x_en == 0:
                pass
            elif gen_pulse_i2x_en == 1:
                if self.pulse1_reg_cmd != 'x':
                    mcu_s.pulse_out(self.pulse1_reg_cmd, self.pulse2_data_cmd)
                pass
            elif gen_pulse_i2x_en == 2:
                if self.pulse1_reg_cmd != 'x':
                    mcu_s.i2c_single_write(
                        self.pulse1_reg_cmd, self.pulse2_data_cmd)

            # if gen_loader_en == 1 or gen_loader_en == 3:
            # set up all the load current
            # if gen_loader_ch_amount >= 1:
            if self.load_ch1 != 'x':
                load_s.chg_out_auto_mode(self.load_ch1, 1, 'on')
            #     pass

            # if gen_loader_ch_amount > 1:
            if self.load_ch2 != 'x':
                load_s.chg_out_auto_mode(self.load_ch2, 2, 'on')
            #     pass

            # if gen_loader_ch_amount > 2:
            if self.load_ch3 != 'x':
                load_s.chg_out_auto_mode(self.load_ch3, 3, 'on')
            #     pass

            # if gen_loader_ch_amount > 3:
            if self.load_ch4 != 'x':
                load_s.chg_out_auto_mode(self.load_ch4, 4, 'on')
            #     pass

            # if gen_loader_en == 2 or gen_loader_en == 3:
            # setup src
            if self.load_src != 'x':
                load_src_s.change_I(self.load_src, 'on')

            # add the vin calibration
            if vin_cal == 1:
                if gen_pwr_ch_amount >= 1:
                    if self.pwr_ch1 != 'x':
                        temp_res = pwr_s.vin_clibrate_singal_met(
                            0, self.pwr_ch1, met_v_s, mcu_s, excel_s)
                if gen_pwr_ch_amount > 1:
                    if self.pwr_ch2 != 'x':
                        temp_res = pwr_s.vin_clibrate_singal_met(
                            6, self.pwr_ch2, met_v_s, mcu_s, excel_s)
                if gen_pwr_ch_amount > 2:
                    if self.pwr_ch3 != 'x':
                        temp_res = pwr_s.vin_clibrate_singal_met(
                            7, self.pwr_ch3, met_v_s, mcu_s, excel_s)

            # # since vin calibration also return the sting of calibration result,
            # # it doesn't a must to measure Vin of each channel again

            # self.res_met_curr = met_i_s.mea_i()
            # mcu_s.relay_ctrl(0)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v1 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # mcu_s.relay_ctrl(6)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v2 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # mcu_s.relay_ctrl(7)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v3 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # mcu_s.relay_ctrl(1)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v4 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # mcu_s.relay_ctrl(2)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v5 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # mcu_s.relay_ctrl(3)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v6 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # mcu_s.relay_ctrl(4)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v7 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # mcu_s.relay_ctrl(5)
            # # time.sleep(excel_s.wait_small)
            # self.res_met_v8 = met_v_s.mea_v()
            # time.sleep(excel_s.wait_small)

            # self.res_load_curr1 = load_s.read_iout(1)
            # self.res_load_curr2 = load_s.read_iout(2)
            # self.res_load_curr3 = load_s.read_iout(3)
            # self.res_load_curr4 = load_s.read_iout(4)
            # self.res_src_curr = load_src_s.read('CURR')
            # # self.res_temp_read = chamber_s.read('temp_mea')

            self.run_veri_add_in_1(ctrl_index=ctrl_ind_1)
            # control index send from the the call of verification and know if to toggle enable or not
            # need to adjust the call of run_verificatio in the main program

            self.data_measured()
            self.data_latch(x_count, self.obj_sim_mode)
            # latch the data to related position

            self.run_veri_add_in_2(ctrl_index=ctrl_ind_2)

            # save the result and also check program exit
            excel_s.excel_save()
            if excel_s.turn_inst_off == 1:
                self.end_of_exp()
                excel_s.excel_save()

            x_count = x_count + 1
            pass

        print('program finished')
        self.extra_file_name_setup()
        self.inst_off()
        self.table_return()
        self.extra_file_name_setup()
        self.end_of_exp()
        pass

    def run_veri_add_in_1(self, ctrl_index=0):
        '''
        add in function for run verification, check related position and able to input control selection
        for different requirement in different application \n
        this add in is before data measure after every condition is ready\n
        send different ctrl_index for different result
        '''
        self.ctrl_ind_1 = ctrl_index
        if ctrl_index == 0:
            # do nothing and it's the general run_verification function
            pass
        elif ctrl_index == 1:
            # high V buck NT50970 series OTP testing

            # action: toggle EN1 to see if temperature is low enough for buck to recover
            # if not in buck OTP, should be able to turn on
            # not to toggle LDO since LDO should be auto recover

            self.mcu_ini.pmic_mode(3)
            time.sleep(0.2)
            self.mcu_ini.pmic_mode(4)
            time.sleep(0.05)

            pass

        elif ctrl_index == 2:
            # high V buck band gap, change to standby mode and measure LDO output drift
            # before the buck measurement
            self.mcu_ini.pmic_mode(3)
            time.sleep(0.2)
            self.data_measured()
            self.data_latch(
                index=self.x_count, test_mode_b=self.obj_sim_mode, other_sheet=self.ex_sh_array[0])
            self.mcu_ini.pmic_mode(4)
            time.sleep(0.05)

        pass

    def run_veri_add_in_2(self, ctrl_index=0):
        '''
        add_in_2 is placed after data measurement,
        ctrl_index: 0 => nothing, 1 => pre-short

        '''

        self.ctrl_ind_2 = ctrl_index
        if ctrl_index == 0:
            # doing nothing
            pass
        elif ctrl_index == 1:
            # for the pre-short operation
            # need to turn off the power supply after the measurement
            # turn off power supply based on the pwr index

            if self.excel_ini.gen_pwr_ch_amount >= 1:
                if self.pwr_ch1 != 'x':
                    self.pwr_ini.chg_out(
                        self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'off')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.excel_ini.gen_pwr_i_set,
                                         self.excel_ini.relay0_ch, 'off')
            if self.excel_ini.gen_pwr_ch_amount > 1:
                if self.pwr_ch2 != 'x':
                    self.pwr_ini.chg_out(
                        self.pwr_ch2, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay6_ch, 'off')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.excel_ini.gen_pwr_i_set,
                                         self.excel_ini.relay6_ch, 'off')
            if self.excel_ini.gen_pwr_ch_amount > 2:
                if self.pwr_ch3 != 'x':
                    self.pwr_ini.chg_out(
                        self.pwr_ch3, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay7_ch, 'off')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.excel_ini.gen_pwr_i_set,
                                         self.excel_ini.relay7_ch, 'off')

            pass
        print(f'end of the run_veri_add_in_2 in index: {ctrl_index}')

        pass

    def set_sheet_name(self, ctrl_sheet_name0, extra_sheet=0, extra_name='_'):
        '''
        choose the sheet name for general testing\n
        extra_sheet = 1, it will create new sheet in extra array, access from excel array\n
        it will also return the related sheet
        extra_name is to add new comment at the end of new sheet name
        '''

        # assign the related sheet of each format gen
        if extra_sheet == 0:
            # this means the sheet is control sheet
            self.ctrl_sheet_name = ctrl_sheet_name0

            # sh_general_test is the sheet can be access from other object
            # load the setting value for instrument
            self.excel_ini.sh_general_test = self.excel_ini.wb.sheets(
                str(self.ctrl_sheet_name))
            self.sh_general_test = self.excel_ini.wb.sheets(
                str(self.ctrl_sheet_name))

            # also include the new sheet setting from each different sheet

            # loading the control values
            # 220926: index of counter need to passed to the excel, so other object or instrument
            # is able to reference

            self.c_test_amount = self.sh_general_test.range('B6').value
            # only record one loop control variable to prevent error, if none element is
            # in the blank, try to filled 0 in the blank

            # start to adjust the the format based on the input settings

            self.new_sheet_name = str(self.sh_general_test.range('B2').value)

            print('sheet name ready')
            # auto update instrument setting after the sheet is assigned
            self.update_inst_settings()
            self.sheet_name_ready = 1

            # clear all the result saving parameter
            self.res_met_curr = 0
            self.res_met_v1 = 0
            self.res_met_v2 = 0
            self.res_met_v3 = 0
            self.res_met_v4 = 0
            self.res_met_v5 = 0
            self.res_met_v6 = 0
            self.res_met_v7 = 0
            self.res_met_v8 = 0
            self.res_load_curr1 = 0
            self.res_load_curr2 = 0
            self.res_load_curr3 = 0
            self.res_load_curr4 = 0
            self.res_src_curr = 0
            self.res_temp_read = 0
            self.pwr_relay0_ioout = 0
            self.pwr_relay6_ioout = 0
            self.pwr_relay7_ioout = 0

            print('contrl sheet assigned')
            # 221220: control also allow to add extra sheet name
            self.sheet_gen(extra_sh=0, extra_name0=extra_name)
            # no need to assign control sheet and return
            # since control sheet is already record to the excel
            extra_sheet = 'this is control sheet'
            pass

        else:
            # this is not control sheet, just copy and return the
            extra_sheet = self.excel_ini.wb.sheets(str(ctrl_sheet_name0))
            extra_sheet = self.sheet_gen(
                extra_sh=extra_sheet, extra_name0=extra_name)

            # transfer extra sheet control variable to the mapped sheet
            # need to copy to result book
            pass

        # give the sheet generation
        # extra_sheet = self.sheet_gen(extra_sh=extra_sheet)
        return extra_sheet
        # here will return the string if not control extra sheet

        pass

    def sheet_gen(self, extra_sh=0, extra_name0=''):

        if self.sheet_name_ready == 0:
            print('no proper sheet name set yet, need to set_sheet_name first')
            print('or maybe need to assign control sheet')
            pass
        else:
            # this function is a must have function to generate the related excel for this verification item
            # this sub must include:
            # 2. generate the result sheet in the result book, and setup the format
            # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
            # 4. not a new file but an add on sheet to the result workbook

            # # copy the rsult sheet to result book
            # self.excel_ini.sh_general_test.copy(self.excel_ini.sh_ref)
            # # assign the sheet to result book
            # self.excel_ini.sh_general_test = self.excel_ini.wb_res.sheets(
            #     str(self.ctrl_sheet_name))
            # 221209: since .copy will return the cpoied sheet, just assign, no need for name
            if extra_sh == 0:
                # this is control sheet, not extra sheet
                self.excel_ini.sh_general_test = self.excel_ini.sh_general_test.copy(
                    self.excel_ini.sh_ref)
                extra_sh = 'this is control sheet'

                # change the sheet name after finished and save into the excel object
                if extra_name0 != '_':
                    # need to add the extra name for control sheet also
                    self.excel_ini.sh_general_test.name = str(
                        self.new_sheet_name) + str(extra_name0)
                else:
                    self.excel_ini.sh_general_test.name = str(
                        self.new_sheet_name)
                self.sh_general_test = self.excel_ini.sh_general_test
            else:
                # for the extra sheet, just copy and return the extra sheet
                extra_sh = extra_sh.copy(self.excel_ini.sh_ref)
                new_ex_sh_name = str(extra_sh.range('B2').value)
                extra_sh.name = new_ex_sh_name + str(extra_name0)
                '''
                extra_name0 need to be assign when using extra sheet to prevent error,
                maximum of extra sheet should be 5 (based on the array)
                '''

                # after create the extra sheet, assign to extra sheet list and add the counter
                self.ex_sh_array[self.extra_count] = extra_sh
                self.extra_count = self.extra_count + 1

                pass
        return extra_sh
        pass

    def extra_sh_clear(self):
        '''
        this function used to clear extra_count, when end of this item or when it's been called
        '''
        # also clear in the table return ~
        self.extra_count = 0

        pass

    def check_blank(self, y_index, x_index):
        # send the checking index in and check or filled in 0 for none
        check_temp = self.sh_general_test.range(y_index, x_index).value
        if check_temp == None:
            check_temp = 'x'
            # return 0 after read and set the blank to 0
            self.sh_general_test.range(y_index, x_index).value = 'x'
            pass

        # return the settings to main program
        # it should be able to be float or string
        # auto adjustment after being used from others
        return check_temp

    def update_inst_settings(self):
        # setup the instrument after sheet selected
        print('instrument parameter update')
        self.pwr_ini.change_I(self.excel_ini.gen_pwr_i_set, 1)
        self.pwr_ini.change_I(self.excel_ini.gen_pwr_i_set, 2)
        self.pwr_ini.change_I(self.excel_ini.gen_pwr_i_set, 3)
        self.iout_r0 = self.excel_ini.gen_pwr_i_set
        self.iout_r6 = self.excel_ini.gen_pwr_i_set
        self.iout_r7 = self.excel_ini.gen_pwr_i_set
        pass

    def inst_off(self):

        print('going to turn off all the instrument')
        self.pwr_ini.chg_out(0, 0, 1, 'off')
        self.pwr_ini.chg_out(0, 0, 2, 'off')
        self.pwr_ini.chg_out(0, 0, 3, 'off')

        self.loader_ini.chg_out(0, 1, 'off')
        self.loader_ini.chg_out(0, 2, 'off')
        self.loader_ini.chg_out(0, 3, 'off')
        self.loader_ini.chg_out(0, 4, 'off')

        self.src_ini.load_off()

        self.chamber_ini.chamber_off()

        pass

    def end_of_exp(self):
        # reset MCU back to default
        self.mcu_ini.back_to_initial()

        print("Grace's one laugh can make me happy one day!")
        time.sleep(0.5)

        self.inst_off()
        self.table_return()
        # self.extra_file_name_setup()
        # enable the change of name from the different function
        self.extra_sh_clear()

        self.excel_ini.ready_to_off = 1

        pass

    def table_return(self):
        # need to recover this sheet: self.excel_ini.sh_ref_table
        self.excel_ini.sh_general_test = self.excel_ini.wb.sheets(
            'general_example')

        # reset sheet choice to wait for next sheet name update
        self.sheet_name_ready = 0

        # reset the count and sheet array of the extra sheet
        self.ex_sh_array = [0, 0, 0, 0, 0]
        self.extra_count = 0

        pass

    def data_measured(self):

        # since vin calibration also return the sting of calibration result,
        # it doesn't a must to measure Vin of each channel again

        self.mcu_ini.relay_ctrl(0)
        # time.sleep(excel_s.wait_small)
        self.res_met_v1 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.mcu_ini.relay_ctrl(6)
        # time.sleep(excel_s.wait_small)
        self.res_met_v2 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.mcu_ini.relay_ctrl(7)
        # time.sleep(excel_s.wait_small)
        self.res_met_v3 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.mcu_ini.relay_ctrl(1)
        # time.sleep(excel_s.wait_small)
        self.res_met_v4 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.mcu_ini.relay_ctrl(2)
        # time.sleep(excel_s.wait_small)
        self.res_met_v5 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.mcu_ini.relay_ctrl(3)
        # time.sleep(excel_s.wait_small)
        self.res_met_v6 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.mcu_ini.relay_ctrl(4)
        # time.sleep(excel_s.wait_small)
        self.res_met_v7 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.mcu_ini.relay_ctrl(5)
        # time.sleep(excel_s.wait_small)
        self.res_met_v8 = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_small)

        self.res_load_curr1 = float(self.loader_ini.read_iout(1))
        self.res_load_curr2 = float(self.loader_ini.read_iout(2))
        self.res_load_curr3 = float(self.loader_ini.read_iout(3))
        self.res_load_curr4 = float(self.loader_ini.read_iout(4))
        self.res_src_curr = self.src_ini.read('CURR')
        # self.res_temp_read = chamber_s.read('temp_mea')
        self.pwr_relay0_ioout = float(self.pwr_ini.read_iout(
            self.excel_ini.relay0_ch))
        self.pwr_relay6_ioout = float(self.pwr_ini.read_iout(
            self.excel_ini.relay6_ch))
        self.pwr_relay7_ioout = float(self.pwr_ini.read_iout(
            self.excel_ini.relay7_ch))

        time.sleep(3 * self.excel_ini.wait_small)
        self.res_met_curr = self.met_i_ini.mea_i()
        time.sleep(self.excel_ini.wait_small)

        pass

    def pwr_iout_set(self, iout_r0=0.1, iout_r6=0.1, iout_r7=0.1):
        '''
        config the different iout of the power supply in general test, default is 0.1A
        need to put after set_sheet_name if going to change current
        '''
        self.iout_r0 = iout_r0
        self.iout_r6 = iout_r6
        self.iout_r7 = iout_r7

        pass

    def gen_pwr_on_off(self, pwr_iout=0):
        '''
        to initial special function for different items
        this example function only for counter and loop
        '''
        dly_tune_ms = 200
        # delay in ms
        dly_set = dly_tune_ms/1000

        dly_measure = 0.2
        # measurement after 200ms of command finished

        # 221222: to record iout or not, need to add the blank in excel
        self.pwr_iout_en = pwr_iout

        # power supply OV and OC protection
        self.pwr_ini.ov_oc_set(self.excel_ini.pre_vin_max,
                               self.excel_ini.pre_imax)

        # power supply channel (channel on setting)
        if self.excel_ini.pre_test_en == 1:
            if self.excel_ini.gen_pwr_ch_amount >= 1:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay0_ch, 'on')
            if self.excel_ini.gen_pwr_ch_amount > 1:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay6_ch, 'on')
            if self.excel_ini.gen_pwr_ch_amount > 2:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay7_ch, 'on')
            print('pre-power on here')
            # turn off the power and load

            print('also turn all load off')

            if self.excel_ini.en_start_up_check == 1:
                self.excel_ini.message_box('press enter if hardware configuration is correct',
                                           'Pre-power on for system test under Vin= ' + str(self.excel_ini.pre_vin) + 'Iin= ' + str(self.excel_ini.pre_sup_iout))
                # msg_res = win32api.MessageBox(
                #     0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(pre_vin) + 'Iin= ' + str(pre_sup_iout))

            # initial settings for power on off is off
            time.sleep(dly_set)
            if self.excel_ini.gen_pwr_ch_amount >= 1:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay0_ch, 'off')
            if self.excel_ini.gen_pwr_ch_amount > 1:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay6_ch, 'off')
            if self.excel_ini.gen_pwr_ch_amount > 2:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay7_ch, 'off')
            # reset MCU to (EN,SW) = (0,0)
            self.mcu_ini.pmic_mode(1)

        x_count = 0
        sub_count = int(self.c_test_amount/8)
        while x_count < self.c_test_amount:
            # load the setting first
            self.data_loaded(x_count)

            # is able to operate high and low temp power on and off
            if self.chamber_target != 'x':
                self.res_temp_read = self.chamber_ini.chamber_set(
                    self.chamber_target)

            # different power on sequence

            # 221221: add other channel control for the power supply
            '''
            for the pwr_on_off program, ch1 is used for Vin,
            other channel shold also able to use for bias or other application
            and it will change follow count_index

            add the same update method like previous, it's able to change with
            channel index setting at main sheet
            '''

            if self.excel_ini.gen_pwr_ch_amount >= 1:
                # this program ch1 is lock for the Vin, control by the
                # other part of program
                pass
            if self.excel_ini.gen_pwr_ch_amount > 1:
                if self.pwr_ch2 != 'x':
                    self.pwr_ini.chg_out(
                        self.pwr_ch2, self.iout_r6, self.excel_ini.relay6_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.iout_r6,
                                         self.excel_ini.relay6_ch, 'off')
            if self.excel_ini.gen_pwr_ch_amount > 2:
                if self.pwr_ch3 != 'x':
                    self.pwr_ini.chg_out(
                        self.pwr_ch3, self.iout_r7, self.excel_ini.relay7_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.iout_r7,
                                         self.excel_ini.relay7_ch, 'off')

            if x_count < 1 * sub_count:
                # sequence1 pwr-EN-SW
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 2 * sub_count:
                # sequence2 pwr-SW-EN
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 3 * sub_count:
                # sequence3 EN-pwr-SW
                self.mcu_ini.pmic_mode(3)
                time.sleep(1.5 * dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 4 * sub_count:
                # sequence4 SW-pwr-EN
                self.mcu_ini.pmic_mode(2)
                time.sleep(1.5 * dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 5 * sub_count:
                # sequence5 EN-SW-pwr
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)
                time.sleep(1.5 * dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')

                pass
            elif x_count < 6 * sub_count:
                # sequence6 SW-EN-pwr
                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)
                time.sleep(1.5 * dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')

                pass
            elif x_count < 7 * sub_count:
                # sequence7 pwr-SW=EN
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                time.sleep(1.5 * dly_set)
                # time.sleep(dly_set)
                # self.mcu_ini.pmic_mode(2)
                self.mcu_ini.pmic_mode(4)
                time.sleep(dly_set)

                pass
            elif x_count < 8 * sub_count:
                # sequence8 SW=EN-pwr
                self.mcu_ini.pmic_mode(4)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                # time.sleep(dly_set)
                # self.mcu_ini.pmic_mode(2)

                pass

            time.sleep(dly_measure)
            self.data_measured()

            # # turn off after data measure
            # self.pwr_ini.chg_out(
            #     self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')
            # # reset MCU to (EN,SW) = (0,0)
            # self.mcu_ini.pmic_mode(1)
            time.sleep(dly_measure)
            self.data_latch(x_count, self.obj_sim_mode)

            # different power off sequence

            if x_count < 1 * sub_count:
                # sequence1 pwr-EN-SW, off: SW-EN-pwr
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(1)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')

                pass
            elif x_count < 2 * sub_count:
                # sequence2 pwr-SW-EN, off: EN-SW-pwr

                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(1)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')

                pass
            elif x_count < 3 * sub_count:
                # sequence3 EN-pwr-SW, off: SW-pwr-EN
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(1)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')

                pass
            elif x_count < 4 * sub_count:
                # sequence4 SW-pwr-EN, off: EN-pwr-SW
                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')
                time.sleep(1.5 * dly_set)
                self.mcu_ini.pmic_mode(1)

                pass
            elif x_count < 5 * sub_count:
                # sequence5 EN-SW-pwr, off: pwr-SW-EN
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')
                time.sleep(1.5 * dly_set)
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(1)

                pass
            elif x_count < 6 * sub_count:
                # sequence6 SW-EN-pwr, off: pwr-EN-SW
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')
                time.sleep(1.5 * dly_set)
                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(1)

                pass
            elif x_count < 7 * sub_count:
                # sequence7 pwr-SW=EN, off: SW=EN-pwr
                self.mcu_ini.pmic_mode(1)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')
                # time.sleep(dly_set)
                # self.mcu_ini.pmic_mode(2)

                pass
            elif x_count < 8 * sub_count:
                # sequence8 SW=EN-pwr, off: pwr-SW=EN
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'off')
                time.sleep(1.5 * dly_set)
                self.mcu_ini.pmic_mode(1)

                # time.sleep(dly_set)
                # self.mcu_ini.pmic_mode(2)
                pass

            # 221215 latch for power off
            time.sleep(dly_measure)
            self.data_measured()
            # here the first is to record power off status
            time.sleep(dly_measure)
            self.data_latch(x_count, test_mode_b=self.obj_sim_mode,
                            other_sheet=self.ex_sh_array[0])

            # save the result and also check program exit
            self.excel_ini.excel_save()
            if self.excel_ini.turn_inst_off == 1:
                self.end_of_exp()
                self.excel_ini.excel_save()

            x_count = x_count + 1

            pass

        print('program finished')
        self.inst_off()
        self.table_return()
        self.extra_file_name_setup(extra_name='_pwr_on_off')
        self.end_of_exp()

        pass

    def pre_short(self, pwr_iout=1, sheet_seq=0, bias_off=0, i_max=1):
        '''
        pre-short testing with chamber
        pwr_iout set to default 1 since iout is needed for the judgement for short
        power supply channel is fixed in pre-short, not follow excel sheeet \n
        channel sequence: 2,1,3 ; 3 is fixed for bias
        if using the original sequence, sheet_seq=1 \n
        there is no Vin calibration in pre-short
        bias_off: control the bias to turn off or not(default not dutn off)
        i_max: default set supply current to the max of pwr supply
        '''
        # index not to turn on the power supply after damage
        # when it become 1, means pre-short damage
        self.pre_short_damage_r0 = 0
        self.pre_short_damage_r6 = 0
        self.pre_short_damage_r7 = 0
        wait_time = 0.5

        if i_max == 1:
            # set the power input to max current and bias to 50mA
            self.pwr_iout_set(iout_r0=3, iout_r6=3, iout_r7=0.05)

        self.pwr_iout_en = pwr_iout

        if sheet_seq == 0:
            # use the pre_short configuration sequence
            # save the temp channel setting and return after pre-short is over
            temp_r0 = self.excel_ini.relay0_ch
            temp_r6 = self.excel_ini.relay6_ch
            temp_r7 = self.excel_ini.relay7_ch

            self.excel_ini.relay0_ch = 2
            self.excel_ini.relay6_ch = 1
            self.excel_ini.relay7_ch = 3

        x_count = 0
        while x_count < self.c_test_amount:
            # load the setting first
            self.data_loaded(x_count)

            # is able to operate high and low temp power on and off
            if self.chamber_target != 'x':
                self.res_temp_read = self.chamber_ini.chamber_set(
                    self.chamber_target)

            # change sequence since bias must power up first
            if self.excel_ini.gen_pwr_ch_amount > 2:
                if self.pwr_ch3 != 'x' and self.pre_short_damage_r7 == 0:
                    self.pwr_ini.chg_out(
                        self.pwr_ch3, self.iout_r7, self.excel_ini.relay7_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.iout_r7,
                                         self.excel_ini.relay7_ch, 'off')

            time.sleep(0.1)
            # make sure bias ready before Vin comes

            if self.excel_ini.gen_pwr_ch_amount >= 1:
                # this program ch1 is lock for the Vin, control by the
                # other part of program

                if self.pwr_ch1 != 'x' and self.pre_short_damage_r0 == 0:
                    self.pwr_ini.chg_out(
                        self.pwr_ch1, self.iout_r0, self.excel_ini.relay0_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.iout_r0,
                                         self.excel_ini.relay0_ch, 'off')

                pass
            if self.excel_ini.gen_pwr_ch_amount > 1:
                if self.pwr_ch2 != 'x' and self.pre_short_damage_r6 == 0:
                    self.pwr_ini.chg_out(
                        self.pwr_ch2, self.iout_r6, self.excel_ini.relay6_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    self.pwr_ini.chg_out(0, self.iout_r6,
                                         self.excel_ini.relay6_ch, 'off')

            time.sleep(wait_time)
            self.data_measured()

            # turn off the power after measurement finished

            if self.excel_ini.gen_pwr_ch_amount >= 1:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay0_ch, 'off')
            if self.excel_ini.gen_pwr_ch_amount > 1:
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay6_ch, 'off')
            if self.excel_ini.gen_pwr_ch_amount > 2 and bias_off == 1:
                # 221230 modify: not to turn off bias is default bias_off == 0
                self.pwr_ini.chg_out(
                    self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout, self.excel_ini.relay7_ch, 'off')

            if float(self.pwr_relay0_ioout) > 1:
                # relay0 already burned
                self.pre_short_damage_r0 = 1

                pass
            if float(self.pwr_relay6_ioout) > 1:
                # relay6 already burned
                self.pre_short_damage_r6 = 1

                pass
            if float(self.pwr_relay7_ioout) > 1:
                # relay7 already burned
                self.pre_short_damage_r7 = 1

                pass

            self.data_latch(x_count)
            # save the result and also check program exit
            self.excel_ini.excel_save()
            if self.excel_ini.turn_inst_off == 1:
                self.end_of_exp()
                self.excel_ini.excel_save()

            x_count = x_count + 1
            pass

        print('program finished')
        self.extra_file_name_setup('pre_short')
        self.inst_off()
        self.table_return()
        self.end_of_exp()

        if sheet_seq == 0:
            # return the sheet setting from temp
            self.excel_ini.relay0_ch = temp_r0
            self.excel_ini.relay6_ch = temp_r6
            self.excel_ini.relay7_ch = temp_r7

        pass

    def flexible_gen_ini(self):
        '''
        to initial special function for different items
        this example function only for counter and loop
        '''

        x_count = 0
        while x_count < self.c_test_amount:
            # load the setting first
            self.data_loaded(x_count)

            # =====

            self.data_measured()

            # =====

            self.data_latch(x_count)
            # save the result and also check program exit
            self.excel_ini.excel_save()
            if self.excel_ini.turn_inst_off == 1:
                self.end_of_exp()
                self.excel_ini.excel_save()

            x_count = x_count + 1
            pass

        print('program finished')
        self.extra_file_name_setup()
        self.inst_off()
        self.table_return()
        self.extra_file_name_setup()
        self.end_of_exp()

        pass

    def data_latch(self, index, test_mode_b=1, other_sheet=0):
        '''
        test_mode_b is the main_off_line used for debug\n
        other_sheet will change the latch to other sheet
        '''
        if other_sheet != 0:
            # change the latch result to other sheet
            sh_temp = self.sh_general_test
            self.sh_general_test = other_sheet
            # return the self.sh setting when out of the function

        # update all the result based on index

        # 221110 send 0 as result for data latch if the instrument is in simulation mode
        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 1).value = lo.atof(self.res_met_v1)
        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 2).value = lo.atof(self.res_met_v2)
        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 3).value = lo.atof(self.res_met_v3)

        if self.met_i_ini.sim_inst == 1 or test_mode_b == 0:
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 4).value = lo.atof(self.res_met_curr)
        else:
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 4).value = 0

        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 5).value = lo.atof(self.res_met_v4)
        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 6).value = lo.atof(self.res_met_v5)
        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 7).value = lo.atof(self.res_met_v6)
        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 8).value = lo.atof(self.res_met_v7)
        self.sh_general_test.range(
            8 + index, 1 + self.excel_ini.gen_col_amount + 9).value = lo.atof(self.res_met_v8)

        if self.loader_ini.sim_inst == 1 or test_mode_b == 0:
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 10).value = lo.atof(self.res_load_curr1)
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 11).value = lo.atof(self.res_load_curr2)
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 12).value = lo.atof(self.res_load_curr3)
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 13).value = lo.atof(self.res_load_curr4)
        else:
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 10).value = 0
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 11).value = 0
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 12).value = 0
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 13).value = 0

        if self.src_ini.sim_inst == 1 or test_mode_b == 0:
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 14).value = lo.atof(self.res_src_curr)
        else:
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 14).value = 0

        if self.chamber_ini.sim_inst == 1 or test_mode_b == 0:
            # temperature return is already float, no need to change
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 15).value = self.res_temp_read
        else:
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 15).value = 0

        if self.pwr_iout_en != 0:
            # latch iout to the result sheet
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 16).value = self.pwr_relay0_ioout
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 17).value = self.pwr_relay6_ioout
            self.sh_general_test.range(
                8 + index, 1 + self.excel_ini.gen_col_amount + 18).value = self.pwr_relay7_ioout

        print('data latch for g finished')

        if other_sheet != 0:
            self.sh_general_test = sh_temp
            # return the self.sh setting when out of the function

        pass

    def data_loaded(self, index):

        # update all the result based on index
        self.pwr_ch1 = self.check_blank(8 + index, 2)
        self.pwr_ch2 = self.check_blank(8 + index, 3)
        self.pwr_ch3 = self.check_blank(8 + index, 4)
        self.load_ch1 = self.check_blank(8 + index, 5)
        self.load_ch2 = self.check_blank(8 + index, 6)
        self.load_ch3 = self.check_blank(8 + index, 7)
        self.load_ch4 = self.check_blank(8 + index, 8)
        self.load_src = self.check_blank(8 + index, 9)
        self.pulse1_reg_cmd = self.check_blank(8 + index, 10)
        self.pulse2_data_cmd = self.check_blank(8 + index, 11)
        self.chamber_target = self.check_blank(8 + index, 12)

        # self.pwr_ch1 = self.sh_general_test.range( 8 + index , 2 ).value
        # self.pwr_ch2 = self.sh_general_test.range( 8 + index , 3 ).value
        # self.pwr_ch3 = self.sh_general_test.range( 8 + index , 4) .value
        # self.load_ch1 = self.sh_general_test.range( 8 + index , 5 ).value
        # self.load_ch2 = self.sh_general_test.range( 8 + index , 6 ).value
        # self.load_ch3 = self.sh_general_test.range( 8 + index , 7 ).value
        # self.load_ch4 = self.sh_general_test.range( 8 + index , 8 ).value
        # self.load_src = self.sh_general_test.range( 8 + index , 9 ).value
        # self.pulse1_reg_cmd = self.sh_general_test.range( 8 + index , 10 ).value
        # self.pulse2_data_cmd = self.sh_general_test.range( 8 + index , 11 ).value
        # self.chamber_target = self.sh_general_test.range( 8 + index , 12 ).value

        print('data loaded for g finished')

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
    pwr_t.sim_inst = 1
    pwr_t.open_inst()
    # initial the object and set to simulation mode
    met_v_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
    met_v_t.sim_inst = 0
    met_v_t.open_inst()
    load_t = inst.chroma_63600(1, 7, 'CCL')
    load_t.sim_inst = 0
    load_t.open_inst()
    met_i_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_i_t.sim_inst = 0
    met_i_t.open_inst()
    src_t = inst.Keth_2440(0, 0, 24, 'off', 'CURR', 15)
    src_t.sim_inst = 0
    src_t.open_inst()
    chamber_t = inst.chamber_su242(25, 15, 'off', -45, 180, 0)
    chamber_t.sim_inst = 0
    chamber_t.open_inst()
    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu_t = mcu.MCU_control(0, 13)
    mcu_t.com_open()

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

    version_select = 4

    if version_select == 0:
        # create one object

        general_t = general_test(
            excel_t, pwr_t, met_v_t, load_t, mcu_t, src_t, met_i_t, chamber_t)
        general_t.set_sheet_name('general_1')
        general_t.sheet_gen()
        general_t.run_verification()
        general_t.table_return()

        general_t.set_sheet_name('general_2')
        general_t.sheet_gen()
        general_t.run_verification()
        general_t.table_return()

        excel_t.end_of_file(0)

    elif version_select == 1:
        #  reduce the sheet_gen and table rerun function in the main

        general_t = general_test(
            excel_t, pwr_t, met_v_t, load_t, mcu_t, src_t, met_i_t, chamber_t)
        general_t.set_sheet_name('general_1')
        # general_t.sheet_gen()
        general_t.run_verification()
        # general_t.table_return()

        general_t.set_sheet_name('general_2')
        # general_t.sheet_gen()
        general_t.run_verification()
        # general_t.table_return()

        excel_t.end_of_file(0)

        pass
    elif version_select == 2:

        general_t = general_test(
            excel_t, pwr_t, met_v_t, load_t, mcu_t, src_t, met_i_t, chamber_t)

        general_t.set_sheet_name('general_2')

        general_t.gen_pwr_on_off()

        excel_t.end_of_file(0)
        pass

    elif version_select == 3:
        # for power on and off testing, need to assign second sheet for
        # record the power off result

        general_t = general_test(
            excel_t, pwr_t, met_v_t, load_t, mcu_t, src_t, met_i_t, chamber_t)

        general_t.set_sheet_name('general_1', 0)
        general_t.set_sheet_name('general_1', 1, '_pwr_off')
        general_t.gen_pwr_on_off()

        excel_t.end_of_file(0)
        pass

    elif version_select == 4:
        # testing for pre-short function
        general_t = general_test(
            excel_t, pwr_t, met_v_t, load_t, mcu_t, src_t, met_i_t, chamber_t)

        general_t.set_sheet_name('gen_pre_short_HT_HV', extra_sheet=0)
        general_t.pre_short()
