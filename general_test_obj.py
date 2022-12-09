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

    pass

    def extra_file_name_setup(self):
        self.excel_ini.extra_file_name = '_general'
        pass

    def run_verification(self, vin_cal=1):
        '''
        run the general testing: default calibrate Vin on
        to disable, change vin_cal to 0
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
                        self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    pwr_s.chg_out(0, self.excel_ini.gen_pwr_i_set,
                                  self.excel_ini.relay0_ch, 'off')
            if gen_pwr_ch_amount > 1:
                if self.pwr_ch2 != 'x':
                    pwr_s.chg_out(
                        self.pwr_ch2, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay6_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    pwr_s.chg_out(0, self.excel_ini.gen_pwr_i_set,
                                  self.excel_ini.relay6_ch, 'off')
            if gen_pwr_ch_amount > 2:
                if self.pwr_ch3 != 'x':
                    pwr_s.chg_out(
                        self.pwr_ch3, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay7_ch, 'on')
                else:
                    # turn off the power if not going to control power
                    pwr_s.chg_out(0, self.excel_ini.gen_pwr_i_set,
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

            self.res_met_curr = met_i_s.mea_i()

            # # since vin calibration also return the sting of calibration result,
            # # it doesn't a must to measure Vin of each channel again

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

            self.data_measured()

            self.data_latch(x_count, self.obj_sim_mode)
            # latch the data to related position

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
        pass

    def set_sheet_name(self, ctrl_sheet_name0):

        # assign the related sheet of each format gen
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

        # give the sheet generation
        self.sheet_gen()

        pass

    def sheet_gen(self):

        if self.sheet_name_ready == 0:
            print('no proper sheet name set yet, need to set_sheet_name first')
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
            self.excel_ini.sh_general_test = self.excel_ini.sh_general_test.copy(self.excel_ini.sh_ref)

            # change the sheet name after finished and save into the excel object
            self.excel_ini.sh_general_test.name = str(self.new_sheet_name)
            self.sh_general_test = self.excel_ini.sh_general_test

            pass

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
        self.extra_file_name_setup()

        self.excel_ini.ready_to_off = 1

        pass

    def table_return(self):
        # need to recover this sheet: self.excel_ini.sh_ref_table
        self.excel_ini.sh_general_test = self.excel_ini.wb.sheets(
            'general_example')

        # reset sheet choice to wait for next sheet name update
        self.sheet_name_ready = 0

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

        self.res_load_curr1 = self.loader_ini.read_iout(1)
        self.res_load_curr2 = self.loader_ini.read_iout(2)
        self.res_load_curr3 = self.loader_ini.read_iout(3)
        self.res_load_curr4 = self.loader_ini.read_iout(4)
        self.res_src_curr = self.src_ini.read('CURR')
        # self.res_temp_read = chamber_s.read('temp_mea')

        pass

    def gen_pwr_on_off(self):
        '''
        to initial special function for different items
        this example function only for counter and loop
        '''
        dly_tune_ms = 0
        # delay in ms
        dly_set = dly_tune_ms/1000

        dly_measure = 0.2
        # measurement after 200ms of command finished

        x_count = 0
        sub_count = int(self.c_test_amount/8)
        while x_count < self.c_test_amount:
            # load the setting first
            self.data_loaded(x_count)

            # is able to operate high and low temp power on and off
            if self.chamber_target != 'x':
                self.res_temp_read = self.chamber_ini.chamber_set(
                    self.chamber_target)

            if x_count < 1 * sub_count:
                # sequence1 pwr-EN-SW
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 2 * sub_count:
                # sequence2 pwr-SW-EN
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 3 * sub_count:
                # sequence3 EN-pwr-SW
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 4 * sub_count:
                # sequence4 SW-pwr-EN
                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)

                pass
            elif x_count < 5 * sub_count:
                # sequence5 EN-SW-pwr
                self.mcu_ini.pmic_mode(3)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')

                pass
            elif x_count < 6 * sub_count:
                # sequence6 SW-EN-pwr
                self.mcu_ini.pmic_mode(2)
                time.sleep(dly_set)
                self.mcu_ini.pmic_mode(4)
                time.sleep(dly_set)
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')

                pass
            elif x_count < 7 * sub_count:
                # sequence7 pwr-SW=EN
                self.pwr_ini.chg_out(
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')
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
                    self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'on')
                # time.sleep(dly_set)
                # self.mcu_ini.pmic_mode(2)

                pass

            time.sleep(dly_measure)
            self.data_measured()

            # turn off after data measure
            self.pwr_ini.chg_out(
                self.pwr_ch1, self.excel_ini.gen_pwr_i_set, self.excel_ini.relay0_ch, 'off')
            # reset MCU to (EN,SW) = (0,0)
            self.mcu_ini.pmic_mode(1)

            self.data_latch(x_count, self.obj_sim_mode)

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

        pass

    def data_latch(self, index, test_mode_b=1):
        '''
        test_mode_b is the main_off_line used for debug
        '''

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

        print('data latch for g finished')

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
    pwr_t.sim_inst = 0
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
    mcu_t = mcu.MCU_control(0, 5)
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

    version_select = 2

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
