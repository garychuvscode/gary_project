# 220903: for the new structure, using object to define each function
# efficiency measurement object


# excel parameter and settings
from csv import excel
import parameter_load_obj as par
# for the jump out window
# # also for the jump out window, same group with win32con
import win32api
from win32con import MB_SYSTEMMODAL
# for the delay function
import time
# include for atof function => transfer string to float
import locale as lo


class eff_mea:

    # this class is used to measure IQ from the DUT, based on the I/O setting and different Vin
    # measure the IQ

    def __init__(self, excel0, pwr0, met_v0, loader_0, mcu0, src0, met_i0, chamber0):

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
        if self.excel_ini.channel_mode == 0:
            self.excel_ini.extra_file_name = '_EL_EFF'
            pass
        elif self.excel_ini.channel_mode == 1:
            self.excel_ini.extra_file_name = '_AVDD_EFF'
            pass
        elif self.excel_ini.channel_mode == 2:
            self.excel_ini.extra_file_name = '_3ch_EFF'
            pass
        pass

    # since there are more than 1 file for efficiency test, need to call file name reset
    def extra_file_name_setup(self):
        if self.excel_ini.channel_mode == 0:
            self.excel_ini.extra_file_name = '_EL_EFF'
            pass
        elif self.excel_ini.channel_mode == 1:
            self.excel_ini.extra_file_name = '_AVDD_EFF'
            pass
        elif self.excel_ini.channel_mode == 2:
            self.excel_ini.extra_file_name = '_3ch_EFF'
            pass
        pass

    def sheet_gen(self):
        # this function is a must have function to generate the related excel for this verification item
        # this sub must include:
        # 2. generate the result sheet in the result book, and setup the format
        # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
        # 4. not a new file but an add on sheet to the result workbook

        # copy the rsult sheet to result book
        self.excel_ini.sh_volt_curr_cmd.copy(self.excel_ini.sh_ref)
        # assign the sheet to result book
        self.excel_ini.sh_volt_curr_cmd = self.excel_ini.wb_res.sheets(
            'V_I_com')
        # copy the rsult sheet to result book
        self.excel_ini.sh_i2c_cmd.copy(self.excel_ini.sh_ref)
        # assign the sheet to result book
        self.excel_ini.sh_i2c_cmd = self.excel_ini.wb_res.sheets('I2C_ctrl')
        # copy the rsult sheet to result book
        self.excel_ini.sh_raw_out.copy(self.excel_ini.sh_ref)
        # assign the sheet to result book
        self.excel_ini.sh_raw_out = self.excel_ini.wb_res.sheets('raw_out')

        # # copy the sheets to new book
        # # for the new sheet generation, located in sheet_gen
        # self.excel_s.sh_main.copy(self.sh_ref_condition)
        # self.sh_result.copy(self.sh_ref)

        pass

    def table_plot(self):
        # this function need to build the plot needed for this verification
        # include the VBA function inside the excel

        pass

    def inst_name_for_eff(self):

        self.excel_ini.inst_name_sheet('PWR1', self.pwr_ini.inst_name())
        self.excel_ini.inst_name_sheet('MET1', self.met_v_ini.inst_name())
        self.excel_ini.inst_name_sheet('MET2', self.met_i_ini.inst_name())
        self.excel_ini.inst_name_sheet('LOAD1', self.loader_ini.inst_name())
        self.excel_ini.inst_name_sheet('LOADSR', self.src_ini.inst_name())
        self.excel_ini.inst_name_sheet('chamber', self.chamber_ini.inst_name())
        pass

    def run_verification(self):
        # this function is to run the main item, for all the instrument control and main loop will be in this sub function
        # for the parameter only loaded to the program, no need to call from boject all the time
        # save to local variable every time call the run_verification program

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

        # sheet needed in the sub
        res_sheet = excel_s.sh_sw_scan

        # specific variable for each verification
        pwr_act_ch = excel_s.pwr_act_ch
        vin_set = excel_s.vin_set
        iin_set = excel_s.Iin_set
        EL_curr = excel_s.EL_curr
        VCI_curr = excel_s.VCI_curr
        loader_ELch = excel_s.loader_ELch
        loader_VCIch = excel_s.loader_VCIch
        c_swire = excel_s.c_swire
        wait_time = excel_s.wait_time
        # wait_small = excel_s.wait_small

        en_start_up_check = excel_s.en_start_up_check
        pre_test_en = excel_s.pre_test_en
        relay0_ch = excel_s.relay0_ch
        loader_ELch = excel_s.loader_ELch
        loader_VCIch = excel_s.loader_VCIch
        source_meter_channel = excel_s.source_meter_channel
        en_chamber_mea = excel_s.eff_chamber_en
        inst_auto_selection = excel_s.auto_inst_ctrl
        chamber_default_tset = excel_s.cham_tset_ini
        c_tempature = excel_s.c_tempature
        sw_i2c_select = excel_s.sw_i2c_select
        channel_mode = excel_s.channel_mode
        c_avdd_pulse = excel_s.c_avdd_pulse
        c_pulse = excel_s.c_pulse
        c_i2c = excel_s.c_i2c
        c_i2c_g = excel_s.c_i2c_g
        c_avdd_load = excel_s.c_avdd_load
        wait_small = excel_s.wait_small
        sheet_arry = excel_s.sheet_arry
        sub_sh_count = excel_s.sub_sh_count
        raw_gap = excel_s.raw_gap
        c_avdd_single = excel_s.c_avdd_single
        c_iload = excel_s.c_iload
        c_vin = excel_s.c_vin

        # move all as much as possible index to the front of sub program and easier to modify
        # and less chance to fail when change index

        # 220911 local variable needed for the verification process
        eff_done = 0
        bypass_measurement_flag = 0

        # 220521 new start of program, infinite loop for the selection
        # choose the eff or the normal instrument control
        while (excel_s.program_exit > 0):
            excel_s.program_exit = excel_s.sh_main.range('B12').value
            excel_s.auto_inst_ctrl = excel_s.sh_main.range('B11').value
            inst_auto_selection = excel_s.auto_inst_ctrl

            # check if the eff_done need to reset before the loop
            # eff_done can only be reset when it's already be 1 to prevent error
            if eff_done == 1:
                excel_s.eff_rerun()
                eff_done = excel_s.eff_done_sh

            if inst_auto_selection == 1:
                # 220824 change the condition to prevent stuck in check instrument when running efficiency testing

                # the setting of only instrument control
                excel_s.check_inst_update()

                # call the update function and refresh the settings
                time.sleep(wait_time)
                # give some delay between each refresh (call subprogram)

                pass
            elif (inst_auto_selection == 0 or inst_auto_selection == 2) and eff_done == 0:
                # 220824 change the condition to prevent stuck in check instrument when running efficiency testing

                # here is the original eff testing program
                # in the deepest loop, will check if auto selection is 0 or 2
                # for 2 it will also call the check instrument and control
                # for 0 is just like the previour efficiency measurement

                # 220521: move the insturment initialization to the front
                # settings need for related sheet of instrument

                # protection and other setting (before channel on)

                # power supply OV and OC protection
                pwr_s.ov_oc_set(pre_vin_max, pre_imax)

                # power supply channel (channel on setting)
                if pre_test_en == 1:
                    pwr_s.chg_out(pre_vin, pre_sup_iout,
                                  relay0_ch, 'on')
                    print('pre-power on here')
                    # turn off the power and load

                    load_s.chg_out(0, loader_ELch, 'off')
                    load_s.chg_out(0, loader_VCIch, 'off')

                    if source_meter_channel == 1 or source_meter_channel == 2:
                        load_src_s.load_off()

                    print('also turn all load off')

                    if en_start_up_check == 1:
                        msg_res = win32api.MessageBox(
                            0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(pre_vin) + 'Iin= ' + str(pre_sup_iout))

                    if en_chamber_mea == 1:
                        # chamber turn on with default setting, using default temperature
                        chamber_s.chamber_set(chamber_default_tset)

                # the power will change from initial state directly, not turn off between transition

                # should not need the extra Vin in the change
                # pwr_s.chg_out(vin1_set, pre_sup_iout, relay0_ch, 'on')

                # loader channel and current
                # default off, will be turn on and off based on the loop control

                # load_s.chg_out(iload1_set, loader_ELch, 'off')
                # # load set for EL-power
                # load_s.chg_out(iload2_set, loader_VCIch, 'off')
                # # load set for AVDD

                time.sleep(wait_time)

                # add the while loop outside of SWIRE or I2C loop
                x_temperature = 0
                count_temperature = c_tempature
                if en_chamber_mea == 0:
                    # cancel the temperature counter if en_chamber is disable
                    count_temperature = 1
                while x_temperature < count_temperature:
                    # update temperature setting every time the loop is start
                    tset_now = excel_s.sh_volt_curr_cmd.range(
                        (3 + x_temperature, 9)).value
                    if en_chamber_mea == 1:
                        chamber_s.chamber_set(tset_now)
                    else:
                        tset_now = 25
                        # the temperature without chamber are all assume to be 25C

                    # efficiency testing program starts from heere

                    # 1st loop is the selection of I2C and SWIRE pulse control

                    # selection for the loop control variable
                    x_sw_i2c = 0
                    c_sw_i2c = 0

                    # selection for the SWIRE(1) or I2C(2)
                    if sw_i2c_select == 1:
                        if channel_mode == 1:
                            # if the setting is avdd only ,need to map the counter to AVDD pulse number
                            c_sw_i2c = c_avdd_pulse
                        elif channel_mode == 0 or channel_mode == 2:
                            c_sw_i2c = c_pulse
                    elif sw_i2c_select == 2:
                        c_sw_i2c = c_i2c
                        # i2c group counter setting
                        c_i2c_group = c_i2c_g

                    # 220823 consider to add the AVDD measurement for both I2C and normal(SWIRE mode)
                    # but maybe no need for the

                    # error handling can be consider in future, but now should be ok to prevent
                    # all the error from knowing the system operating rule and bug

                    # # error handling for counter = 0 after data refresh
                    # if c_sw_i2c == 0 :
                    #     c_sw_i2c = 1
                    #     c_single = 1

                    while x_sw_i2c < c_sw_i2c:
                        # need to set up specific SWIRE pulse setting (2-pulse version) or give the I2C command at this stage
                        # before the stage of next loop
                        # loaded the SWIRE or I2C command from the control sheet
                        # send it out from MCU to testing EVM board

                        if sw_i2c_select == 1:
                            # SWIRE control loop
                            # setup the related 2 pulse
                            if channel_mode == 1:
                                # if the setting is avdd only ,need to map the pulse to AVDD pulse
                                pulse1 = excel_s.sh_volt_curr_cmd.range(
                                    (3 + x_sw_i2c, 8)).value
                                pulse2 = pulse1
                                extra_file_name = '_t' + \
                                    str(tset_now) + 'C_' + 'SWIRE_AVDD_'
                            elif channel_mode == 0 or channel_mode == 2:
                                # for the SWIRE pulse, need 2 pulse for ELVDD and ELVSS
                                pulse1 = excel_s.sh_volt_curr_cmd.range(
                                    (3 + x_sw_i2c, 5)).value
                                pulse2 = excel_s.sh_volt_curr_cmd.range(
                                    (3 + x_sw_i2c, 6)).value
                                extra_file_name = '_t' + \
                                    str(tset_now) + 'C_' + 'SWIRE_EL_'
                            print('pulse1: ' + str(pulse1) +
                                  '; and pulse2: ' + str(pulse2) + 'under mode ' + str(channel_mode))
                            # send the pulse out through MCU
                            mcu_s.pulse_out(pulse1, pulse2)

                            # call the build file to build new file to save result
                            # for the SWIRE pulse control
                            extra_file_name = extra_file_name + \
                                str(int(pulse1)) + '_' + str(int(pulse2))
                            excel_s.sw_i2c_status = str(
                                int(pulse1)) + '_' + str(int(pulse2))
                            # build_file(str(extra_file_name))
                            # pro_status_str = 'file built'
                            # program_status(pro_status_str)

                        elif sw_i2c_select == 2:
                            # setup i2C group counter
                            x_i2c_group = 0

                            # modify the i2c str before the loop to get all change (register:data)
                            # initial the extra_file_name before the loop start
                            extra_file_name = '_t' + \
                                str(tset_now) + 'C_' + 'i2c'
                            while x_i2c_group < c_i2c_group:
                                # I2C control loop
                                # set up the i2c related data
                                reg_i2c = excel_s.sh_i2c_cmd.range(
                                    (3 + c_i2c_group * x_sw_i2c + x_i2c_group, 2)).value
                                data_i2c = excel_s.sh_i2c_cmd.range(
                                    (3 + c_i2c_group * x_sw_i2c + x_i2c_group, 3)).value
                                print('register: ' + reg_i2c)
                                print('data: ' + data_i2c)

                                mcu_s.i2c_single_write(reg_i2c, data_i2c)

                                # after write to the MCU, update the extra name for the file
                                extra_file_name = extra_file_name + '_' + \
                                    str(reg_i2c) + '-' + str(data_i2c)
                                excel_s.sw_i2c_status = str(
                                    reg_i2c) + '-' + str(data_i2c)
                                x_i2c_group = x_i2c_group + 1

                            # extra_file_name = 'SWIRE_' + str(pulse1) + '_' + str(pulse2)

                        # extra_file_name here is only the loacal of run_verification
                        # it's call detail_name in the excel object
                        excel_s.build_file(str(extra_file_name))

                        # since sub sheet count is update after build file
                        sub_sh_count = excel_s.sub_sh_count
                        pro_status_str = 'file built'
                        excel_s.program_status(pro_status_str)
                        print(excel_s.sh_main)
                        print('')
                        # 220823 change the file build code to outer loop and save the few lines before

                        # finished the swire or i2c command and ready to enter next loop

                        # # call the build file to build new file to save result
                        # build_file(str(x_sw_i2c))

                        x_avdd = 0
                        # counter avdd current setting
                        # when in channel_mode 0 or 1, there is not option of AVDD current
                        # need to add the error handling of AVDD current counter in the loop
                        # need to set the aVDD current counter to 1 to pevent of overflow in the sheet array
                        c_avdd = 0
                        if channel_mode == 0 or channel_mode == 1:
                            # for the 1 or 2 channel operation, only used c_avdd once
                            # there are no extra loop at this stage
                            c_avdd = 1
                        else:
                            # this is 3-channel operatoion
                            c_avdd = c_avdd_load

                        while x_avdd < c_avdd:

                            # define AVDD current
                            curr_avdd = excel_s.sh_volt_curr_cmd.range(
                                (3 + x_avdd, 4)).value
                            print('AVDD current is set to: ' +
                                  str(curr_avdd) + ' A')
                            if channel_mode == 2 or channel_mode == 0:
                                excel_s.i_avdd_status = str(curr_avdd)
                                # 20220509 => need to add the (EN, SW) mode setting to the MCU status
                                # since the i2c don't have the mode selection function, don't care
                                # about the PMIC status, need to be config in the register command in
                                # i2c mode

                                pmic_mode = 4
                                mcu_s.pmic_mode(pmic_mode)
                                print('MCU mode is set to (EN, SW) = (1, 1)')
                            else:
                                excel_s.i_avdd_status = 0
                                # since the i2c don't have the mode selection function, don't care
                                # about the PMIC status, need to be config in the register command in
                                # i2c mode

                                # 0511: to have better discharge result for PMIC, add the shut down mode between
                                # and then start from AOD mode only
                                pmic_mode = 1
                                mcu_s.pmic_mode(pmic_mode)
                                print('PMIC shut down for EL power discharge')
                                # turn off the PMIC and wait for other channel to discharge
                                time.sleep(wait_small)

                                pmic_mode = 3
                                # set the PMIC to AOD mode
                                # and update the MCU write commanad
                                mcu_s.pmic_mode(pmic_mode)
                                print('MCU mode is set to (EN, SW) = (1, 0)')

                            pro_status_str = 'AVDD current : ' + str(curr_avdd)
                            excel_s.program_status(pro_status_str)
                            # please note that AVDD current need to set to 0 if not using 3-ch mode

                            # after generate the related file, should be able to have array for active raw and sheet
                            # and the mapping sheet name is change with AVDD current, so it's in AVDD_current loop

                            excel_s.sheet_active = excel_s.wb_res.sheets(
                                sheet_arry[sub_sh_count * x_avdd])
                            excel_s.eff_temp = str(
                                sheet_arry[sub_sh_count * x_avdd])
                            # add the string save sheet name for the usage of plot
                            # 220911 assign the sheet variable to local variable after change index
                            sheet_active = excel_s.sheet_active

                            excel_s.raw_active = excel_s.wb_res.sheets(
                                sheet_arry[sub_sh_count * x_avdd + 1])
                            excel_s.raw_temp = str(
                                str(sheet_arry[sub_sh_count * x_avdd + 1]))
                            # add the string save sheet name for the usage of plot
                            # 220911 assign the sheet variable to local variable after change index
                            raw_active = excel_s.raw_active

                            excel_s.vout_p_active = excel_s.wb_res.sheets(
                                sheet_arry[sub_sh_count * x_avdd + 2])
                            excel_s.pos_temp = str(
                                sheet_arry[sub_sh_count * x_avdd + 2])
                            # add the string save sheet name for the usage of plot
                            # vout_p_active => is the active sheet (sheet boject)
                            # pos_temp is th string of sheet name which used as the reference
                            # for plotting function
                            # 220825 AVDD and ELVDD is sharing the same sheet assignment since
                            # it won't be record at the same time
                            # 220911 assign the sheet variable to local variable after change index
                            vout_p_active = excel_s.vout_p_active

                            if channel_mode == 1:
                                excel_s.vout_p_pre_active = excel_s.wb_res.sheets(
                                    sheet_arry[sub_sh_count * x_avdd + 3])
                                excel_s.pos_pre_temp = str(
                                    sheet_arry[sub_sh_count * x_avdd + 3])
                                # 220911 assign the sheet variable to local variable after change index
                                vout_p_pre_active = excel_s.vout_p_pre_active

                            elif channel_mode == 0 or channel_mode == 2:
                                excel_s.vout_p_pre_active = excel_s.wb_res.sheets(
                                    sheet_arry[sub_sh_count * x_avdd + 4])
                                excel_s.pos_pre_temp = str(
                                    sheet_arry[sub_sh_count * x_avdd + 4])
                                # 220911 assign the sheet variable to local variable after change index
                                vout_p_pre_active = excel_s.vout_p_pre_active

                            # 220825 add the sheet mapping for Vout and Von

                            if channel_mode == 0 or channel_mode == 2:
                                # only assign the sheet if there are negative output used in the measurement
                                excel_s.vout_n_active = excel_s.wb_res.sheets(
                                    sheet_arry[sub_sh_count * x_avdd + 3])
                                excel_s.neg_temp = str(
                                    sheet_arry[sub_sh_count * x_avdd + 3])
                                # add the string save sheet name for the usage of plot
                                # 220911 assign the sheet variable to local variable after change index
                                vout_n_active = excel_s.vout_n_active

                                excel_s.vout_n_pre_active = excel_s.wb_res.sheets(
                                    sheet_arry[sub_sh_count * x_avdd + 5])
                                excel_s.neg_pre_temp = str(
                                    sheet_arry[sub_sh_count * x_avdd + 5])
                                # 220911 assign the sheet variable to local variable after change index
                                vout_n_pre_active = excel_s.vout_n_pre_active

                            # ====================
                            #  this portion seems not a must have portion in the system
                            # if channel_mode == 0 or channel_mode == 2 :
                            #     # EL only or 3-ch operation
                            #     sheet_active = wb_res.sheets(
                            #         sheet_arry[sub_sh_count * x_avdd])
                            #     raw_active = wb_res.sheets(
                            #         sheet_arry[sub_sh_count * x_avdd + 1])
                            #     vout_p_active = wb_res.sheets(
                            #         sheet_arry[sub_sh_count * x_avdd + 2])
                            #     vout_p_active = wb_res.sheets(
                            #         sheet_arry[sub_sh_count * x_avdd + 3])
                            # else :
                            #     # AVDD only
                            #     sheet_active = wb_res.sheets(
                            #         sheet_arry[sub_sh_count * x_avdd])
                            #     raw_active = wb_res.sheets(
                            #         sheet_arry[sub_sh_count * x_avdd + 1])
                            #     vout_p_active = wb_res.sheets(
                            #         sheet_arry[sub_sh_count * x_avdd + 2])
                            # ====================

                            # not to turn load on here, Vin haven't change yet, only update the
                            # AVDD load current here

                            # if sim_real == 1 :
                            #     load_s.chg_out(curr_avdd, loader_VCIch, 'on')
                            #     # turn the load of AVDD on when after load the current
                            # else:
                            #     print('finished set the current and turn load on')
                            #     # input()

                            # Vin loop start point
                            # counter for the Vin setting

                            x_vin = 0
                            while x_vin < c_vin:

                                v_target = excel_s.sh_volt_curr_cmd.range(
                                    (3 + x_vin, 2)).value
                                pro_status_str = 'Vin:' + str(v_target)
                                excel_s.vin_status = str(v_target)
                                excel_s.program_status(pro_status_str)
                                # update the target Vin to the program status

                                # add the related Vin(ideal) setting at the result sheet
                                sheet_active.range(
                                    (24, 3 + x_vin)).value = v_target

                                # 220325: regulation sheet also need to have current index
                                if channel_mode == 0 or channel_mode == 2:
                                    # ELVDD and ELVSS have regulation sheet
                                    vout_p_active.range(
                                        (24, 3 + x_vin)).value = v_target
                                    vout_n_active.range(
                                        (24, 3 + x_vin)).value = v_target
                                    # 220825 add the V setting for Vop and Von sheet
                                    vout_p_pre_active.range(
                                        (24, 3 + x_vin)).value = v_target
                                    vout_n_pre_active.range(
                                        (24, 3 + x_vin)).value = v_target
                                else:
                                    # AVDD have regulation sheet
                                    vout_p_active.range(
                                        (24, 3 + x_vin)).value = v_target
                                    # 220825 add the V setting for Vop and Von sheet
                                    vout_p_pre_active.range(
                                        (24, 3 + x_vin)).value = v_target

                                # for the raw data sheet index
                                raw_active.range(
                                    (11 + raw_gap * x_vin, 2)).value = 'Vin'
                                raw_active.range(
                                    (12 + raw_gap * x_vin, 2)).value = 'Iin'
                                raw_active.range(
                                    (13 + raw_gap * x_vin, 2)).value = 'ELVDD'
                                raw_active.range(
                                    (14 + raw_gap * x_vin, 2)).value = 'ELVSS'
                                raw_active.range(
                                    (15 + raw_gap * x_vin, 2)).value = 'I_EL'
                                raw_active.range(
                                    (16 + raw_gap * x_vin, 2)).value = 'AVDD'
                                raw_active.range((17 + raw_gap * x_vin, 2)
                                                 ).value = 'I_AVDD'
                                raw_active.range(
                                    (18 + raw_gap * x_vin, 2)).value = 'Eff'
                                raw_active.range(
                                    (19 + raw_gap * x_vin, 2)).value = 'VOP'
                                raw_active.range(
                                    (20 + raw_gap * x_vin, 2)).value = 'VON'

                                # adjust the vin voltage
                                pwr_s.chg_out(v_target, pre_imax,
                                              relay0_ch, 'on')
                                time.sleep(wait_small)
                                print('vin setting change: ' + str(v_target))

                                # current loop start point
                                # counter of current setting
                                x_iload = 0
                                # here means I_EL

                                # selection for current counter based on the channel setting
                                # when testing for AVDD single channel, need to use AVDD current mapping column
                                # for the loop counter
                                c_load_curr = 0
                                if channel_mode == 1:
                                    # channel mode: 0-EL, 1-AVDD, 2-EL+AVDD
                                    c_load_curr = c_avdd_single
                                else:
                                    c_load_curr = c_iload

                                # for each loop of iload, need to define a new offset
                                value_i_offset1 = 0
                                value_i_offset2 = 0
                                while x_iload < c_load_curr:
                                    # because different mode need to measure different channel,
                                    # rebuild the selection part for items below ... wait for continue (0325)

                                    if channel_mode == 0 or channel_mode == 2:
                                        # for EL only or 3-ch, load target is for EL power
                                        iload_target = excel_s.sh_volt_curr_cmd.range(
                                            (3 + x_iload, 3)).value
                                    else:
                                        # if only AVDD efficiency, assign the load target to AVDD_1ch column
                                        iload_target = excel_s.sh_volt_curr_cmd.range(
                                            (3 + x_iload, 7)).value

                                    # all i_load target must start from 0 mA for calibration

                                    pro_status_str = 'setting iload_target current'
                                    excel_s.i_el_status = str(iload_target)
                                    print(pro_status_str)
                                    excel_s.program_status(pro_status_str)

                                    if x_vin == 0:
                                        sheet_active.range(
                                            (25 + x_iload, 2)).value = iload_target
                                        # only process once for the current index at the result sheet
                                        # 220325: regulation sheet also need to have current index
                                        if channel_mode == 0 or channel_mode == 2:
                                            # ELVDD and ELVSS have regulation sheet
                                            vout_p_active.range(
                                                (25 + x_iload, 2)).value = iload_target
                                            vout_n_active.range(
                                                (25 + x_iload, 2)).value = iload_target
                                            # 220825 add the Vop and Von
                                            vout_p_pre_active.range(
                                                (25 + x_iload, 2)).value = iload_target
                                            vout_n_pre_active.range(
                                                (25 + x_iload, 2)).value = iload_target
                                        else:
                                            # AVDD have regulation sheet
                                            vout_p_active.range(
                                                (25 + x_iload, 2)).value = iload_target
                                            # 220825 add the Vop and Von
                                            vout_p_pre_active.range(
                                                (25 + x_iload, 2)).value = iload_target
                                        pro_status_str = 'give index to regulation sheet'
                                        print(pro_status_str)
                                        excel_s.program_status(pro_status_str)

                                    # setting up for the relay channel

                                    # uart_cmd_str = (chr(mcu_mode_8_bit_IO) +
                                    #                 mcu_cmd_arry[int(meter_ch_ctrl)])
                                    # print(uart_cmd_str)
                                    # mcu_com.write(uart_cmd_str)
                                    # 220328: reset the relay channel to initial state(Vin stage)
                                    meter_ch_ctrl = 0
                                    mcu_s.relay_ctrl(meter_ch_ctrl)
                                    time.sleep(wait_small)
                                    # input()
                                    # finished adjust relay channel

                                    # note that the meter_ch_ctrl will change through the measurement channel change
                                    # so the the start point of meter should be Vin(calibration)

                                    # if sim_real == 1:
                                    # setting up for the instrument control
                                    time.sleep(wait_small)

                                    # initialization the temp saving parameter for the efficiency calculation
                                    # clear result for each round of the current loop

                                    value_eff = 0
                                    if x_iload > 0:
                                        # turn the load on to setting i_load (when x > 0)

                                        # 20220429 method to assign source meter for the PMIC measurement
                                        # only single AVDD efficiency need to use source meter at the AVDD
                                        # other is mainly for EL-power
                                        # make the selection easier in this applicaation

                                        if channel_mode == 2:
                                            # 3 channel mode, only EL powere will have source meter
                                            # selection only at the iload_target but not curr_avdd

                                            # turn on AVDD load channel, must be chroma
                                            load_s.chg_out(curr_avdd,
                                                           loader_VCIch, 'on')

                                            # choose the source meter or the chroma load
                                            if source_meter_channel == 1:
                                                load_src_s.change_I(
                                                    iload_target, 'on')
                                                # or you can use 'keep' to replace 'on' and call the load turn at other point

                                            else:
                                                # not using source meter for EL, it's chroma loader
                                                load_s.chg_out(
                                                    iload_target, loader_ELch, 'on')

                                        elif channel_mode == 0:
                                            # only turn the EL on
                                            # choose the source meter or the chroma load
                                            if source_meter_channel == 1:
                                                load_src_s.change_I(
                                                    iload_target, 'on')
                                                # or you can use 'keep' to replace 'on' and call the load turn at other point

                                            else:
                                                # not using source meter for EL, it's chroma loader
                                                load_s.chg_out(
                                                    iload_target, loader_ELch, 'on')
                                        elif channel_mode == 1:
                                            # only turn the AVDD on
                                            # choose the source meter or the chroma load
                                            if source_meter_channel == 2:
                                                load_src_s.change_I(
                                                    iload_target, 'on')
                                                # or you can use 'keep' to replace 'on' and call the load turn at other point

                                            else:
                                                # not using source meter for EL, it's chroma loader
                                                load_s.chg_out(
                                                    iload_target, loader_VCIch, 'on')

                                    # need to set muc_sim to 1 before using calibration
                                    v_res_temp = pwr_s.vin_clibrate_singal_met(
                                        0, v_target, met_v_s, mcu_s, excel_s)
                                    # record Vin after calibration finished
                                    time.sleep(wait_time)
                                    excel_s.data_latch(
                                        'vin', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)

                                    # all the channel change with original sequence
                                    # but bypass result to 'NA' with related mode of settings
                                    # 0 is bypass and 1 is enable

                                    # =====
                                    meter_ch_ctrl = meter_ch_ctrl + 1
                                    # change relay to AVDD
                                    if channel_mode == 3:
                                        # bypass AVDD measurement result if only check EL efficiency
                                        bypass_measurement_flag = 1
                                    else:
                                        # only change relay and measurement voltage if needed
                                        mcu_s.relay_ctrl(meter_ch_ctrl)
                                        v_res_temp = met_v_s.mea_v()
                                        time.sleep(wait_time)
                                    # for the data latch, bypass flag will decide to use the result or not
                                    excel_s.data_latch(
                                        'avdd', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)

                                    # =====

                                    # =====
                                    meter_ch_ctrl = meter_ch_ctrl + 1
                                    # change relay to ELVDD
                                    if channel_mode == 3:
                                        # bypass ELVDD measurement result (AVDD only)
                                        bypass_measurement_flag = 1
                                    else:
                                        # only change relay and measurement voltage if needed
                                        mcu_s.relay_ctrl(meter_ch_ctrl)
                                        v_res_temp = met_v_s.mea_v()
                                        time.sleep(wait_time)
                                    # for the data latch, bypass flag will decide to use the result or not
                                    excel_s.data_latch(
                                        'elvdd', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)

                                    # =====

                                    # =====
                                    meter_ch_ctrl = meter_ch_ctrl + 1
                                    # change relay to ELVSS
                                    if channel_mode == 3:
                                        # bypass ELVSS measurement result (AVDD only)
                                        bypass_measurement_flag = 1
                                    else:
                                        # only change relay and measurement voltage if needed
                                        mcu_s.relay_ctrl(meter_ch_ctrl)
                                        v_res_temp = met_v_s.mea_v()
                                        time.sleep(wait_time)
                                    # for the data latch, bypass flag will decide to use the result or not
                                    excel_s.data_latch(
                                        'elvss', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)

                                    # =====

                                    # =====
                                    meter_ch_ctrl = meter_ch_ctrl + 1
                                    # change relay to VOP
                                    if channel_mode == 3:
                                        # record in all condition
                                        bypass_measurement_flag = 1
                                    else:
                                        # only change relay and measurement voltage if needed
                                        mcu_s.relay_ctrl(meter_ch_ctrl)
                                        v_res_temp = met_v_s.mea_v()
                                        time.sleep(wait_time)
                                    # for the data latch, bypass flag will decide to use the result or not
                                    excel_s.data_latch(
                                        'vop', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)

                                    # =====

                                    # =====
                                    meter_ch_ctrl = meter_ch_ctrl + 1
                                    # change relay to VON
                                    if channel_mode == 3:
                                        # bypass VON measurement result (AVDD only)
                                        bypass_measurement_flag = 1
                                    else:
                                        # only change relay and measurement voltage if needed
                                        mcu_s.relay_ctrl(meter_ch_ctrl)
                                        v_res_temp = met_v_s.mea_v()
                                        time.sleep(wait_time)
                                    # for the data latch, bypass flag will decide to use the result or not
                                    excel_s.data_latch(
                                        'von', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)

                                    # =====

                                    # # this is the power supply method, not the meter method
                                    # v_res_temp = pwr_s.read_iout()
                                    # # for current reading, need to remove A in the end of string
                                    # v_res_temp = v_res_temp.replace('A', '')
                                    # # this part can also consider to move to the next ints_pkg file
                                    # # can help to improve the complexity

                                    # adjust the Iin measurement from power supply to meter
                                    v_res_temp = met_i_s.mea_i()
                                    excel_s.data_latch(
                                        'iin', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)
                                    time.sleep(wait_time)
                                    # different mode need different Iout => read all Iout but only keep the good one

                                    # 20220429 read I function: read all the channel, but choose to latch or not,
                                    # get the I read result from both chroma channel, choose different way to latch data
                                    # based on the channel mode from control sheet
                                    if source_meter_channel == 0:
                                        v_res_temp = load_s.read_iout(
                                            loader_ELch)
                                        if channel_mode == 0 or channel_mode == 2:
                                            excel_s.data_latch(
                                                'i_el', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)
                                            if channel_mode == 0:
                                                # give the i_avdd blank to 0 for result
                                                excel_s.data_latch(
                                                    'i_avdd', str(value_i_offset2))
                                                # 0511 to preven calibration settings cause error
                                                # pass the calibration parameter into data_latch to
                                                # cancel the result adjustment
                                        v_res_temp = load_s.read_iout(
                                            loader_VCIch)
                                        if channel_mode == 1 or channel_mode == 2:
                                            excel_s.data_latch(
                                                'i_avdd', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)
                                            if channel_mode == 1:
                                                # give the i_el blank to 0 for result
                                                excel_s.data_latch('i_el', '0')
                                                excel_s.data_latch(
                                                    'i_el', str(value_i_offset1))
                                                # 0511 to preven calibration settings cause error
                                                # pass the calibration parameter into data_latch to
                                                # cancel the result adjustment
                                    elif source_meter_channel == 1:
                                        v_res_temp = load_src_s.read(
                                            'CURR')
                                        excel_s.data_latch(
                                            'i_el', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)
                                        if channel_mode == 2 or channel_mode == 0:
                                            # if now is 3-channel mode, also need to latch the current at AVDD
                                            v_res_temp = load_s.read_iout(
                                                loader_VCIch)
                                            excel_s.data_latch(
                                                'i_avdd', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)
                                    elif source_meter_channel == 2:
                                        # since we already assume there is no EL power measurement
                                        # for using source meter for AVDD, EL power will be latch to 0
                                        v_res_temp = load_src_s.read(
                                            'CURR')
                                        excel_s.data_latch(
                                            'i_avdd', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)
                                        # give the i_el blank to 0 for result
                                        excel_s.data_latch('i_el', '0')

                                    # # 220911 cancel the lader offset calibration from main and added to loader
                                    # # ==== loader offset configuration
                                    # # add the calibration factor to iload_target
                                    # if x_iload == 0:

                                    #     if loader_cal_mode == 2:
                                    #         # when the calibration mode is 2, use no load case for the offset adjustment
                                    #         if channel_mode == 0 or channel_mode == 2:
                                    #             value_i_offset1 = value_iel
                                    #             if channel_mode == 2:
                                    #                 value_i_offset2 = value_iavdd
                                    #         elif channel_mode == 1:
                                    #             value_i_offset2 = value_iavdd
                                    #     elif loader_cal_mode == 1:
                                    #         # when the calibration mode is 1, constant offset adjustment
                                    #         if channel_mode == 0 or channel_mode == 2:
                                    #             value_i_offset1 = loader_cal_off1
                                    #             if channel_mode == 2:
                                    #                 value_i_offset2 = loader_cal_off2
                                    #         elif channel_mode == 1:
                                    #             value_i_offset2 = loader_cal_off2
                                    #     else:
                                    #         # the case don't need offset
                                    #         value_i_offset1 = 0
                                    #         value_i_offset2 = 0

                                    #     if source_meter_channel == 1:
                                    #         # if the channel 1=> EL power, 2=> AVDD, 0=> not use source meter
                                    #         value_i_offset1 = 0
                                    #         # remove the offset if using source meter
                                    #     elif source_meter_channel == 2:
                                    #         value_i_offset2 = 0

                                    # # ==== loader offset end point
                                    print(excel_s.sh_main)
                                    pass

                                    # 220329: for the new version of inst_pkg_a, add the A remove function in
                                    # the sub-program operation
                                    # # for current reading, need to remove A in the end of string
                                    # v_res_temp = v_res_temp.replace('A', '')
                                    # excel_s.data_latch('iout', v_res_temp, x_vin, x_iload, value_i_offset1, value_i_offset2)
                                    time.sleep(wait_time)
                                    # record iin and iout for from the source and load
                                    # meter channel shift, when the cycle end, reset the channel
                                    # selection for the next round
                                    meter_ch_ctrl = 0
                                    mcu_s.relay_ctrl(meter_ch_ctrl)
                                    # after setting the meter channel, also give MCU command back to initial
                                    # get ready for the next cycle

                                    # measure Vout
                                    # measure Iin (power supply or meter)
                                    # mrasure I out (loader or meter)

                                    # to prevent overspec of the votlage input, need to change the Vin back to
                                    # intital target before loading release
                                    pwr_s.chg_out(v_target, pre_sup_iout,
                                                  relay0_ch, 'on')
                                    print('V_target return to normal')

                                    # release loading
                                    # turn the load off after measurement
                                    if channel_mode == 2:
                                        # load_s.chg_out(curr_avdd, loader_VCIch, 'off')
                                        # load_s.chg_out(iload_target, loader_ELch, 'off')
                                        load_s.chg_out(
                                            0, loader_VCIch, 'on')
                                        load_s.chg_out(
                                            0, loader_ELch, 'on')
                                    elif channel_mode == 0:
                                        # only turn the EL on
                                        # load_s.chg_out(iload_target, loader_ELch, 'off')
                                        if source_meter_channel == 1 or source_meter_channel == 2:
                                            # load_src_s.load_off()
                                            # change to turn off at each voltage cycle for loadr and source meter
                                            load_src_s.change_I(0, 'on')
                                        else:
                                            load_s.chg_out(
                                                0, loader_ELch, 'on')
                                    elif channel_mode == 1:
                                        # only turn the AVDD on
                                        # load_s.chg_out(iload_target, loader_VCIch, 'off')
                                        if source_meter_channel == 1 or source_meter_channel == 2:
                                            # load_src_s.load_off()
                                            # change to turn off at each voltage cycle for loadr and source meter
                                            load_src_s.change_I(0, 'on')
                                        else:
                                            load_s.chg_out(
                                                0, loader_VCIch, 'on')

                                    # 20220429 since release the load and set to turn off is ok,
                                    # no specific setting for the chroma load selection here
                                    # source meter is also turn off directly

                                    # 220824 to prevent the wrong turning off of the E-load when using source meter for sngle channel operation
                                    # need to change this source meter turn off into loader turn off
                                    # if source_meter_channel == 1 or source_meter_channel == 2:
                                    #     # load_src_s.load_off()
                                    #     # change to turn off at each voltage cycle for loadr and source meter
                                    #     load_src_s.change_I(0, 'on')

                                    # after the result fix in the data saving excel, calculate the efficiency
                                    value_eff = excel_s.eff_calculated()
                                    excel_s.data_latch('eff', str(
                                        value_eff), x_vin, x_iload, value_i_offset1, value_i_offset2)
                                    # latch and locked down the efficiency result

                                    # prepare for the next round

                                    # save the result after each counter finished

                                    # before x_iload increase, need to checck inst status based on insturment control
                                    # 220522 add the check inst sub program
                                    if inst_auto_selection == 2:
                                        excel_s.check_inst_update()
                                    if excel_s.program_exit == 0:
                                        # exit the program
                                        break
                                    # check every time goes to the loop

                                    x_iload = x_iload + 1
                                    # end of the 4th loop

                                # wb_res.save(result_book_trace)
                                excel_s.excel_save()
                                if excel_s.program_exit == 0:
                                    # exit the program
                                    break
                                x_vin = x_vin + 1
                                # end of the 3rd loop

                            # wb_res.save(result_book_trace)
                            # save the file one time after each avdd load current is finished

                            # make the plot for each AVDD current change, since there are one group of the efficiency and regulation data
                            # start from EL only but need to check all 3 mode of the operation
                            v_cnt = c_vin
                            i_cnt = c_load_curr

                            # # decide from the sheet need to plot chart
                            # book_n = str(excel_s.result_book_trace) + '.xlsx'

                            # 220524: for the instrument operation, prevent the issue of change windoww,
                            # need to have message remind user to release control of excel
                            # so there will not be the error from the plot
                            # remind the operation can keep going after the plot is finished
                            if excel_s.en_plot_waring == 1:
                                msg_res = win32api.MessageBox(
                                    0, 'Release the control of excel, change the window to auto file now, and press enter, remind again when the plot is finished', 'Plot request from python')

                            # plot for efficiency
                            sheet_n = excel_s.eff_temp
                            excel_s.plot_single_sheet(v_cnt, i_cnt, sheet_n)

                            if channel_mode == 0 or channel_mode == 2:

                                # plot for ELVDD
                                sheet_n = excel_s.pos_temp
                                excel_s.plot_single_sheet(
                                    v_cnt, i_cnt, sheet_n)

                                # plot for ELVSS
                                sheet_n = excel_s.neg_temp
                                excel_s.plot_single_sheet(
                                    v_cnt, i_cnt, sheet_n)

                                # plot for Vout
                                sheet_n = excel_s.pos_pre_temp
                                excel_s.plot_single_sheet(
                                    v_cnt, i_cnt, sheet_n)

                                # plot for Von
                                sheet_n = excel_s.neg_pre_temp
                                excel_s.plot_single_sheet(
                                    v_cnt, i_cnt, sheet_n)

                            else:
                                # plot for AVDD
                                sheet_n = excel_s.pos_temp
                                excel_s.plot_single_sheet(
                                    v_cnt, i_cnt, sheet_n)

                                # plot for Vout
                                sheet_n = excel_s.pos_pre_temp
                                excel_s.plot_single_sheet(
                                    v_cnt, i_cnt, sheet_n)

                            excel_s.excel_save()
                            # wb_res.save(result_book_trace)
                            print(excel_s.sh_main)
                            print('')
                            # save the result after plot is finished

                            # the plot request is finished and jump another window to remind
                            if excel_s.en_plot_waring == 1:
                                msg_res = win32api.MessageBox(
                                    0, 'You can start to operate the computer again', 'Plot request finished ')

                            if excel_s.program_exit == 0:
                                # exit the program
                                break

                            # change the sheet name if eff file is set to one file
                            if excel_s.eff_single_file == 1:
                                excel_s.sheet_adj_for_eff(x_avdd)
                                pass

                            x_avdd = x_avdd + 1
                            # end of the 2nd loop
                        if excel_s.program_exit == 0:
                            # exit the program
                            break

                        # 220824 add the exit action for save and turn off the excel
                        # 220911 all the turn off control and setting is in end_of_file
                        if excel_s.eff_single_file == 0:
                            # multi-file is ok with efficiency output
                            excel_s.end_of_file(0)
                            # to preven the issue of re-run (same file opening and crash)

                            # to connect with the new program flow settings
                            excel_s.open_result_book()
                            self.sheet_gen()
                            self.inst_name_for_eff()
                            self.extra_file_name_setup()
                            pass
                        else:
                            # need to be single file for all the items
                            excel_s.excel_save()

                        x_sw_i2c = x_sw_i2c + 1
                        # end of the SWIRE/I2C loop

                    x_temperature = x_temperature + 1
                    # end of the temperature loop

                # turn off the load and source after the loop is finished

                # turn off the power and load
                pwr_s.chg_out(0, pre_imax, relay0_ch, 'off')
                load_s.chg_out(0, loader_ELch, 'off')
                load_s.chg_out(0, loader_VCIch, 'off')

                if source_meter_channel == 1 or source_meter_channel == 2:
                    load_src_s.load_off()

                eff_done = 1
                # eff done is set after one cycle of efficiency measurement
                # efficiency measurement will need to reset
                # change the eff_done status on the main sheet
                eff_done_sh = 1
                excel_s.sh_main.range('C13').value = 1
                print(excel_s.sh_main)
                print('')

                # the eff test is finished, add single function call settings
            elif inst_auto_selection == 3:
                # IQ mode for the operation
                # wait for the re-factory result
                pass

        print('this is the end of simulation mode ')
        print('close of instrument is control by main, not single object')
        print('finsihed and goodbye')

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
    excel1 = par.excel_parameter('obj_main')
    # ======== only for object programming

    # open the result book for saving result
    excel1.open_result_book()

    # change simulation mode delay (in second)
    excel1.sim_mode_delay(0.02, 0.01)
    inst.wait_time = 0.01
    inst.wait_samll = 0.01

    # and the different verification method can be call below

    # create one file
    eff_test = eff_mea(excel1, pwr_t, met_v_t, load_t,
                       mcu_t, src_t, met_i_t, chamber_t)

    # generate(or copy) the needed sheet to the result book
    eff_test.sheet_gen()

    # start the testing
    eff_test.run_verification()

    # remember that this is only call by main, not by  object
    excel1.end_of_file(0)

    print('end of the SWIRE scan object testing program')
