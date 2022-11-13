# 220914: this file used to crate the object of inctrument control interface in excel
# pwr supply will have polling vin calibration
# other function wait for future
import time


class instrument_scan ():

    def __init__(self, excel0, pwr0, met_v0, loader_0, mcu0, src0, met_i0, chamber0):
        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst
            # excel parameter and settings
            import parameter_load_obj as par
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

        # object simulation mode
        # set to 1 in default
        self.sim_real = 1

        pass

    def check_inst_update(self):
        # first load the refesh parameter from the control sheet
        self.excel_ini.check_refresh()

        pwr_refresh = self.excel_ini.insctl_pwr_refresh
        load_refresh = self.excel_ini.insctl_load_refresh
        met1_refresh = self.excel_ini.insctl_met1_refresh
        met2_refresh = self.excel_ini.insctl_met2_refresh
        src_refresh = self.excel_ini.insctl_src_refresh

        # 0522 add this part to keep update the control parameter
        # when every time call the sub program
        self.excel_ini.program_exit = self.excel_ini.sh_main.range('B12').value
        # program exit will check every loop finished and break until stop if set ot 0

        # self.excel_ini.inst_auto_selection = self.excel_ini.sh_main.range('B11').value
        # this object is only used for self auto testing, no need to disable due to
        # conflict like eff_inst, so just set this to 1 directly,
        # reference this variable to B11 if going to mapped with auto control

        # pwr refresh
        if pwr_refresh == 0:
            # refresh every time enter the program, run refresh process
            self.excel_ini.para_update_pwr()
            # transfer the on and off command to related string

            # channel 1 update and control
            if self.excel_ini.insctl_pwr_outs_ch1 == 0:
                # 220514: need to transfer command 0 and 1 to the command need for the instrument
                temp_status = 'off'
            elif self.excel_ini.insctl_pwr_outs_ch1 == 1:
                temp_status = 'on'
            if self.sim_real == 1:
                # only give command to the instrument when real mode enable
                self.pwr_ini.chg_out(self.excel_ini.insctl_pwr_vset_ch1,
                                     self.excel_ini.insctl_pwr_iset_ch1, self.excel_ini.relay6_ch, temp_status)
                print('PWR change ch1, vset: ' + str(self.excel_ini.insctl_pwr_vset_ch1) + '; iset: ' +
                      str(self.excel_ini.insctl_pwr_iset_ch1) + '; status: ' + str(temp_status))
                if self.excel_ini.insctl_pwr_outs_ch1 == 1:
                    # if the power channel is turned on, need the Vin calibration
                    if self.excel_ini.insctl_pwr_calibration == 1:
                        # need to have vin calibration
                        v_res_pwr_ch1 = self.pwr_ini.vin_clibrate_singal_met(
                            6, self.excel_ini.insctl_pwr_vset_ch1, self.met_v_ini, self.mcu_ini, self.excel_ini)
                    self.excel_ini.pwr_v_ch1_status = v_res_pwr_ch1
                    self.excel_ini.pwr_i_ch1_status = self.pwr_ini.read_iout(
                        self.excel_ini.relay6_ch)
                    pass
                else:
                    # set the status of ouptut voltage to 0
                    # for the unused channel of power supply
                    self.excel_ini.pwr_v_ch1_status = 0
                    self.excel_ini.pwr_i_ch1_status = 0

                self.excel_ini.pwr_o_ch1_status = temp_status
                # updatet the on and off status here

                pass
            elif self.sim_real == 0:
                # run the testing of simulation mode of update
                # just check the variable from the terminal
                print('now is sim mode of pwr update')
                print('PWR change ch1, vset: ' + str(self.excel_ini.insctl_pwr_vset_ch1) + '; iset: ' +
                      str(self.excel_ini.insctl_pwr_iset_ch1) + '; status: ' + str(temp_status))
                # also set the status to equal to input settings
                self.excel_ini.pwr_v_ch1_status = self.excel_ini.insctl_pwr_vset_ch1
                self.excel_ini.pwr_i_ch1_status = self.excel_ini.insctl_pwr_iset_ch1
                self.excel_ini.pwr_o_ch1_status = temp_status

                pass
            # channel 1 end ======

            # channel 2 update and control
            if self.excel_ini.insctl_pwr_outs_ch2 == 0:
                # 220514: need to transfer command 0 and 1 to the command need for the instrument
                temp_status = 'off'
            elif self.excel_ini.insctl_pwr_outs_ch2 == 1:
                temp_status = 'on'
            if self.sim_real == 1:
                # only give command to the instrument when real mode enable
                self.pwr_ini.chg_out(self.excel_ini.insctl_pwr_vset_ch2,
                                     self.excel_ini.insctl_pwr_iset_ch2, self.excel_ini.relay7_ch, temp_status)
                print('PWR change ch2, vset: ' + str(self.excel_ini.insctl_pwr_vset_ch2) + '; iset: ' +
                      str(self.excel_ini.insctl_pwr_iset_ch2) + '; status: ' + str(temp_status))
                if self.excel_ini.insctl_pwr_outs_ch2 == 1:
                    # if the power channel is turned on, need the Vin calibration
                    if self.excel_ini.insctl_pwr_calibration == 1:
                        # need to have vin calibration
                        v_res_pwr_ch2 = self.pwr_ini.vin_clibrate_singal_met(
                            7, self.excel_ini.insctl_pwr_vset_ch2, self.met_v_ini, self.mcu_ini, self.excel_ini)
                    self.excel_ini.pwr_v_ch2_status = v_res_pwr_ch2
                    self.excel_ini.pwr_i_ch2_status = self.pwr_ini.read_iout(
                        self.excel_ini.relay7_ch)
                    pass
                else:
                    # set the status of ouptut voltage to 0
                    # for the unused channel of power supply
                    self.excel_ini.pwr_v_ch2_status = 0
                    self.excel_ini.pwr_i_ch2_status = 0

                self.excel_ini.pwr_o_ch2_status = temp_status
                # updatet the on and off status here

                pass
            elif self.sim_real == 0:
                # run the testing of simulation mode of update
                # just check the variable from the terminal
                print('now is sim mode of pwr update')
                print('PWR change ch2, vset: ' + str(self.excel_ini.insctl_pwr_vset_ch2) + '; iset: ' +
                      str(self.excel_ini.insctl_pwr_iset_ch2) + '; status: ' + str(temp_status))
                # also set the status to equal to input settings
                self.excel_ini.pwr_v_ch2_status = self.excel_ini.insctl_pwr_vset_ch2
                self.excel_ini.pwr_i_ch2_status = self.excel_ini.insctl_pwr_iset_ch2
                self.excel_ini.pwr_o_ch2_status = temp_status

                pass
            # channel 2 end ======

            # channel 3 update and control
            # ch3 is used by auto test, bypass if auto test is already enable
            if self.excel_ini.inst_auto_selection == 1:
                self.excel_ini.pwr_auto_ch3_status = 1
                if self.excel_ini.insctl_pwr_outs_ch3 == 0:
                    # 220514: need to transfer command 0 and 1 to the command need for the instrument
                    temp_status = 'off'
                elif self.excel_ini.insctl_pwr_outs_ch3 == 1:
                    temp_status = 'on'
                if self.sim_real == 1:
                    # only give command to the instrument when real mode enable
                    self.pwr_ini.chg_out(self.excel_ini.insctl_pwr_vset_ch3,
                                         self.excel_ini.insctl_pwr_iset_ch3, self.excel_ini.relay0_ch, temp_status)
                    print('PWR change ch3, vset: ' + str(self.excel_ini.insctl_pwr_vset_ch3) + '; iset: ' +
                          str(self.excel_ini.insctl_pwr_iset_ch3) + '; status: ' + str(temp_status))
                    if self.excel_ini.insctl_pwr_outs_ch3 == 1:
                        # if the power channel is turned on, need the Vin calibration
                        if self.excel_ini.insctl_pwr_calibration == 1:
                            # need to have vin calibration
                            v_res_temp = self.pwr_ini.vin_clibrate_singal_met(
                                0, self.excel_ini.insctl_pwr_vset_ch3, self.met_v_ini, self.mcu_ini, self.excel_ini)
                        self.excel_ini.pwr_v_ch3_status = v_res_temp
                        self.excel_ini.pwr_i_ch3_status = self.pwr_ini.read_iout(
                            self.excel_ini.relay0_ch)
                        pass
                    else:
                        # set the status of ouptut voltage to 0
                        # for the unused channel of power supply
                        self.excel_ini.pwr_v_ch3_status = 0
                        self.excel_ini.pwr_i_ch3_status = 0

                    self.excel_ini.pwr_o_ch3_status = temp_status
                    # updatet the on and off status here

                    pass
                elif self.sim_real == 0:
                    # run the testing of simulation mode of update
                    # just check the variable from the terminal
                    print('now is sim mode of pwr update')
                    print('PWR change ch3, vset: ' + str(self.excel_ini.insctl_pwr_vset_ch3) + '; iset: ' +
                          str(self.excel_ini.insctl_pwr_iset_ch3) + '; status: ' + str(temp_status))
                    # also set the status to equal to input settings
                    self.excel_ini.pwr_v_ch3_status = self.excel_ini.insctl_pwr_vset_ch3
                    self.excel_ini.pwr_i_ch3_status = self.excel_ini.insctl_pwr_iset_ch3
                    self.excel_ini.pwr_o_ch3_status = temp_status

                    pass
                # channel 3 end ======
                pass
            else:
                self.excel_ini.pwr_auto_ch3_status = 0
                # channel 3 is used from auto test, set to 0 and turn red

            # bypass the serial function first, since there are no serial needed yet
            # only need to change the serial function when line transient control is added to the
            # auto testing, otherwise, no need for the setial finction
            # send the serial control command here for the PWR, but may need to add the control stringin the
            # instpkg file
            # 220516

            # after the setting of power finished, vall the function update status
            self.excel_ini.status_update_pwr()

            pass
        elif pwr_refresh == 1:
            # at the latch mode, pass directly
            pass
        elif pwr_refresh == 2:
            # follow the general refresh settings
            # decide to skip this function first, build in the future if needed
            # 220514: to complicated, skip too many different refresh fuction, and build other part first
            # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
            # other will check after finished, if real a must have function or not
            pass

        pass

        # loader refresh time
        if load_refresh == 0:
            # refresh every time enter the program, run refresh process
            self.excel_ini.para_update_load()
            # transfer the on and off command to related string
            # not

            # ch1 update and control
            # byapss the ch1 and ch2 loader if auto testing is enable
            if self.excel_ini.inst_auto_selection == 1:
                if self.excel_ini.insctl_load_outs_ch1 == 0:
                    # 220514: need to transfer command 0 and 1 to the command need for the instrument
                    temp_status = 'off'
                elif self.excel_ini.insctl_load_outs_ch1 == 1:
                    temp_status = 'on'
                if self.sim_real == 1:
                    # only give command to the instrument when real mode enable
                    # change the mode first to fit the change of current
                    self.loader_ini.chg_mode(
                        1, self.excel_ini.insctl_load_mset_ch1)
                    # change the output current setting after the mode set change
                    self.loader_ini.chg_out(
                        self.excel_ini.insctl_load_iset_ch1, 1, temp_status)
                    print('load change ch1, iset: ' + str(self.excel_ini.insctl_load_iset_ch1) + '; status: ' +
                          str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch1))

                    self.excel_ini.load_i_ch1_status = self.loader_ini.read_iout(
                        1)
                    # read the load current and update to the status
                    self.excel_ini.load_m_ch1_status = self.loader_ini.mode_o[0]
                    # check the final mode setting and update to the status

                    pass
                elif self.sim_real == 0:
                    # run the testing of simulation mode of update
                    # just check the variable from the terminal
                    print('now is sim mode of lpwroad update')
                    print('load change ch1, iset: ' + str(self.excel_ini.insctl_load_iset_ch1) + '; status: ' +
                          str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch1))
                    self.excel_ini.load_i_ch1_status = self.excel_ini.insctl_load_iset_ch1
                    self.excel_ini.load_m_ch1_status = self.excel_ini.insctl_load_outs_ch1
                    pass

                self.excel_ini.load_o_ch1_status = temp_status
                # update the final output status to the status variable
                pass
            else:
                self.excel_ini.load_auto_ch1_status = 0
                # channel 1 is used from auto test, set to 0 and turn red

            # ch1 end ======

            # ch2 update and control
            # byapss the ch1 and ch2 loader if auto testing is enable
            if self.excel_ini.inst_auto_selection == 1:
                if self.excel_ini.insctl_load_outs_ch2 == 0:
                    # 220514: need to transfer command 0 and 1 to the command need for the instrument
                    temp_status = 'off'
                elif self.excel_ini.insctl_load_outs_ch2 == 1:
                    temp_status = 'on'
                if self.sim_real == 1:
                    # only give command to the instrument when real mode enable
                    # change the mode first to fit the change of current
                    self.loader_ini.chg_mode(
                        2, self.excel_ini.insctl_load_mset_ch2)
                    # change the output current setting after the mode set change
                    self.loader_ini.chg_out(
                        self.excel_ini.insctl_load_iset_ch2, 2, temp_status)
                    print('load change ch2, iset: ' + str(self.excel_ini.insctl_load_iset_ch2) + '; status: ' +
                          str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch2))

                    self.excel_ini.load_i_ch2_status = self.loader_ini.read_iout(
                        2)
                    # read the load current and update to the status
                    self.excel_ini.load_m_ch2_status = self.loader_ini.mode_o[1]
                    # check the final mode setting and update to the status

                    pass
                elif self.sim_real == 0:
                    # run the testing of simulation mode of update
                    # just check the variable from the terminal
                    print('now is sim mode of load update')
                    print('load change ch2, iset: ' + str(self.excel_ini.insctl_load_iset_ch2) + '; status: ' +
                          str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch2))

                    self.excel_ini.load_i_ch2_status = self.excel_ini.insctl_load_iset_ch2
                    self.excel_ini.load_m_ch2_status = self.excel_ini.insctl_load_outs_ch2

                    pass

                self.excel_ini.load_o_ch2_status = temp_status
                # update the final output status to the status variable

                pass
            else:
                self.excel_ini.load_auto_ch2_status = 0
                # channel 1 is used from auto test, set to 0 and turn red

            # ch2 end ======

            # ch3 update and control
            if self.excel_ini.insctl_load_outs_ch3 == 0:
                # 220514: need to transfer command 0 and 1 to the command need for the instrument
                temp_status = 'off'
            elif self.excel_ini.insctl_load_outs_ch3 == 1:
                temp_status = 'on'
            if self.sim_real == 1:
                # only give command to the instrument when real mode enable
                # change the mode first to fit the change of current
                self.loader_ini.chg_mode(
                    3, self.excel_ini.insctl_load_mset_ch3)
                # change the output current setting after the mode set change
                self.loader_ini.chg_out(
                    self.excel_ini.insctl_load_iset_ch3, 3, temp_status)
                print('load change ch3, iset: ' + str(self.excel_ini.insctl_load_iset_ch3) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch3))

                self.excel_ini.load_i_ch3_status = self.loader_ini.read_iout(3)
                # read the load current and update to the status
                self.excel_ini.load_m_ch3_status = self.loader_ini.mode_o[2]
                # check the final mode setting and update to the status

                pass
            elif self.sim_real == 0:
                # run the testing of simulation mode of update
                # just check the variable from the terminal
                print('now is sim mode of load update')
                print('load change ch3, iset: ' + str(self.excel_ini.insctl_load_iset_ch3) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch3))

                self.excel_ini.load_i_ch3_status = self.excel_ini.insctl_load_iset_ch3
                self.excel_ini.load_m_ch3_status = self.excel_ini.insctl_load_outs_ch3

                pass

            self.excel_ini.load_o_ch3_status = temp_status
            # update the final output status to the status variable

            # ch3 end ======

            # ch4 update and control
            if self.excel_ini.insctl_load_outs_ch4 == 0:
                # 220514: need to transfer command 0 and 1 to the command need for the instrument
                temp_status = 'off'
            elif self.excel_ini.insctl_load_outs_ch4 == 1:
                temp_status = 'on'
            if self.sim_real == 1:
                # only give command to the instrument when real mode enable
                # change the mode first to fit the change of current
                self.loader_ini.chg_mode(
                    4, self.excel_ini.insctl_load_mset_ch4)
                # change the output current setting after the mode set change
                self.loader_ini.chg_out(
                    self.excel_ini.insctl_load_iset_ch4, 4, temp_status)
                print('load change ch4, iset: ' + str(self.excel_ini.insctl_load_iset_ch4) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch4))

                self.excel_ini.load_i_ch4_status = self.loader_ini.read_iout(4)
                # read the load current and update to the status
                self.excel_ini.load_m_ch4_status = self.loader_ini.mode_o[3]
                # check the final mode setting and update to the status

                pass
            elif self.sim_real == 0:
                # run the testing of simulation mode of update
                # just check the variable from the terminal
                print('now is sim mode of load update')
                print('load change ch4, iset: ' + str(self.excel_ini.insctl_load_iset_ch4) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(self.excel_ini.insctl_load_mset_ch4))

                self.excel_ini.load_i_ch4_status = self.excel_ini.insctl_load_iset_ch4
                self.excel_ini.load_m_ch4_status = self.excel_ini.insctl_load_outs_ch4

                pass

            self.excel_ini.load_o_ch4_status = temp_status
            # update the final output status to the status variable

            # ch4 end ======

            pass
        elif load_refresh == 1:
            # at the latch mode, pass directly
            pass
        elif load_refresh == 2:
            # follow the general refresh settings
            # decide to skip this function first, build in the future if needed
            # 220514: to complicated, skip too many different refresh fuction, and build other part first
            # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
            # other will check after finished, if real a must have function or not
            pass

        self.excel_ini.status_update_load()
        # change the status after the parameter update finished

        # meter1 refresh time
        # add inst_auto selection => when only inst active meter control
        if met1_refresh == 0:
            self.excel_ini.para_update_met1()
            # update the other parameter after refresh checked
            # then start the other parameter update later
            if self.excel_ini.inst_auto_selection == 1:
                # update meter connection to 1 if available
                self.excel_ini.met1_connection_status = 1

                if self.excel_ini.insctl_met1_mset == 0:
                    # enter the voltage measurement mode
                    if self.sim_real == 1:
                        self.excel_ini.met1_v_mea_status = self.met_v_ini.mea_v2(
                            0, self.excel_ini.insctl_met1_leve)
                        time.sleep(self.excel_ini.wait_small)
                        self.excel_ini.met1_i_mea_status = 'NA'
                        self.excel_ini.met1_mode_status = 'votlage'
                        self.excel_ini.met1_level_status = self.excel_ini.insctl_met1_leve
                        pass
                    else:
                        print('meter 1 measure voltage')
                        self.excel_ini.met1_v_mea_status = 'sim_volt1'
                        pass
                    pass

                elif self.excel_ini.insctl_met1_mset == 1:
                    # enter the current measurement mode
                    if self.sim_real == 1:

                        self.excel_ini.met1_i_mea_status = self.met_v_ini.mea_i2(
                            0, self.excel_ini.insctl_met1_leve)
                        time.sleep(self.excel_ini.wait_small)
                        self.excel_ini.met1_v_mea_status = 'NA'
                        self.excel_ini.met1_mode_status = 'current'
                        self.excel_ini.met1_level_status = self.excel_ini.insctl_met1_leve
                        pass
                    else:
                        print('meter 1 measure current')
                        self.excel_ini.met1_i_mea_status = 'sim_curr1'
                        pass
                    pass

                pass
            else:
                # meter channel is decide by auto control
                self.excel_ini.met1_connection_status = 0
                # 0523 other keep in the initial value
        elif met1_refresh == 1:
            # at the latch mode, pass directly
            pass
        elif met1_refresh == 2:
            # follow the general refresh settings
            # decide to skip this function first, build in the future if needed
            # 220514: to complicated, skip too many different refresh fuction, and build other part first
            # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
            # other will check after finished, if real a must have function or not
            pass

        self.excel_ini.status_update_met1()

        # meter2 refresh time
        # add inst_auto selection => when only inst active meter control
        if met2_refresh == 0:
            self.excel_ini.para_update_met2()
            # update the other parameter after refresh checked
            # then start the other parameter update later
            if self.excel_ini.inst_auto_selection == 1:
                self.excel_ini.met2_connection_status = 1

                if self.excel_ini.insctl_met2_mset == 0:
                    # enter the voltage measurement mode
                    if self.sim_real == 1:
                        self.excel_ini.met2_v_mea_status = self.met_i_ini.mea_v2(
                            0, self.excel_ini.insctl_met2_leve)
                        time.sleep(self.excel_ini.wait_small)
                        self.excel_ini.met2_i_mea_status = 'NA'
                        self.excel_ini.met2_mode_status = 'votlage'
                        self.excel_ini.met2_level_status = self.excel_ini.insctl_met2_leve
                        pass
                    else:
                        print('meter 1 measure voltage')
                        self.excel_ini.met2_v_mea_status = 'sim_volt2'
                        pass
                    pass

                elif self.excel_ini.insctl_met2_mset == 1:
                    # enter the current measurement mode
                    if self.sim_real == 1:

                        self.excel_ini.met2_i_mea_status = self.met_i_ini.mea_i2(
                            0, self.excel_ini.insctl_met2_leve)
                        time.sleep(self.excel_ini.wait_small)
                        self.excel_ini.met2_v_mea_status = 'NA'
                        self.excel_ini.met2_mode_status = 'current'
                        self.excel_ini.met2_level_status = self.excel_ini.insctl_met2_leve
                        pass
                    else:
                        print('meter 1 measure current')
                        self.excel_ini.met2_i_mea_status = 'sim_curr2'
                        pass
                    pass
                pass
            else:
                # meter channel is decide by auto control
                # 0523 other keep in the initial value
                self.excel_ini.met2_connection_status = 0

        elif met2_refresh == 1:
            # at the latch mode, pass directly
            pass
        elif met2_refresh == 2:
            # follow the general refresh settings
            # decide to skip this function first, build in the future if needed
            # 220514: to complicated, skip too many different refresh fuction, and build other part first
            # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
            # other will check after finished, if real a must have function or not
            pass

        self.excel_ini.status_update_met2()

        #  source meter refresh time
        # add inst_auto selection => when only inst active source meter control
        if src_refresh == 0:
            self.excel_ini.para_update_src()
            # update the other parameter after refresh checked
            # then start the other parameter update later
            if self.excel_ini.inst_auto_selection == 1:
                self.excel_ini.src_connection_status = 1
                # change the status to string
                if self.excel_ini.insctl_src_outs == 1:
                    self.excel_ini.insctl_src_outs = 'on'
                elif self.excel_ini.insctl_src_outs == 0:
                    self.excel_ini.insctl_src_outs = 'off'

                if self.excel_ini.source_meter_channel == 1 or self.excel_ini.source_meter_channel == 2:

                    if self.excel_ini.insctl_src_mset == 0:
                        # enter the current source mode
                        self.excel_ini.insctl_src_mset = 'CURR'
                        if self.sim_real == 1:

                            if self.src_ini.source_type_o != self.excel_ini.insctl_src_mset:
                                # the only difference is need to turn off load before change mode
                                self.src_ini.load_off()

                            # if the last output state of source meter is same in current source
                            # no need to change type, just change output
                            self.src_ini.change_type(
                                self.excel_ini.insctl_src_mset, self.excel_ini.insctl_src_cset)
                            # if the type doesn't change in the source meter, only change the clamp settings
                            self.src_ini.change_I(self.excel_ini.insctl_src_leve,
                                                  self.excel_ini.insctl_src_outs)
                            # after sending the command to instrument, read the instruemnt status can know
                            # if the command is allowable, command clamp setting should be build in
                            # inst_pkg side, not the main program side

                            # read status and update to the status output blank in inst control sheet
                            self.excel_ini.src_mode_status = self.src_ini.source_type_o
                            self.excel_ini.src_clamp_status = self.src_ini.clamp_VI_o
                            self.excel_ini.src_level_status = self.src_ini.iset_o
                            self.excel_ini.src_v_mea_status = self.src_ini.read(
                                'VOLT')
                            self.excel_ini.src_i_mea_status = self.src_ini.read(
                                'CURR')
                            self.excel_ini.src_o_status = self.src_ini.state_o

                            pass
                        else:
                            # read status and update to the status output blank in inst control sheet
                            self.excel_ini.src_mode_status = self.excel_ini.insctl_src_mset
                            self.excel_ini.src_clamp_status = self.excel_ini.insctl_src_cset
                            self.excel_ini.src_level_status = self.excel_ini.insctl_src_leve
                            self.excel_ini.src_v_mea_status = "self.src_ini.read('VOLT')"
                            self.excel_ini.src_i_mea_status = "self.src_ini.read('CURR')"
                            self.excel_ini.src_o_status = self.excel_ini.insctl_src_outs
                            pass
                        pass

                    elif self.excel_ini.insctl_src_mset == 1:
                        # enter the voltage source mode
                        self.excel_ini.insctl_src_mset = 'VOLT'
                        if self.sim_real == 1:
                            if self.src_ini.state_o != self.excel_ini.insctl_src_mset:
                                # the only difference is need to turn off load before change mode
                                self.src_ini.load_off()

                            # if the last output state of source meter is same in current source
                            # no need to change type, just change output
                            self.src_ini.change_type(
                                self.excel_ini.insctl_src_mset, self.excel_ini.insctl_src_cset)
                            # if the type doesn't change in the source meter, only change the clamp settings
                            self.src_ini.change_V(self.excel_ini.insctl_src_leve,
                                                  self.excel_ini.insctl_src_outs)
                            # after sending the command to instrument, read the instruemnt status can know
                            # if the command is allowable, command clamp setting should be build in
                            # inst_pkg side, not the main program side

                            # read status and update to the status output blank in inst control sheet
                            self.excel_ini.src_mode_status = self.src_ini.source_type_o
                            self.excel_ini.src_clamp_status = self.src_ini.clamp_VI_o
                            self.excel_ini.src_level_status = self.src_ini.iset_o
                            self.excel_ini.src_v_mea_status = self.src_ini.read(
                                'VOLT')
                            self.excel_ini.src_i_mea_status = self.src_ini.read(
                                'CURR')
                            self.excel_ini.src_o_status = self.src_ini.state_o

                        else:
                            # read status and update to the status output blank in inst control sheet
                            self.excel_ini.src_mode_status = self.excel_ini.insctl_src_mset
                            self.excel_ini.src_clamp_status = self.excel_ini.insctl_src_cset
                            self.excel_ini.src_level_status = self.excel_ini.insctl_src_leve
                            self.excel_ini.src_v_mea_status = "self.src_ini.read('VOLT')"
                            self.excel_ini.src_i_mea_status = "self.src_ini.read('CURR')"
                            self.excel_ini.src_o_status = self.excel_ini.insctl_src_outs
                            pass
                        pass
                    else:
                        # since it's not set to voltage or current source
                        # no action for this case
                        pass
                    pass
                pass
            else:
                # set SRC status to 0 if is't control by auto testing
                self.excel_ini.src_connection_status = 0
                # 0523 other keep in the initial value
                pass
        elif src_refresh == 1:
            # at the latch mode, pass directly
            pass
        elif src_refresh == 2:
            # follow the general refresh settings
            # decide to skip this function first, build in the future if needed
            # 220514: to complicated, skip too many different refresh fuction, and build other part first
            # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
            # other will check after finished, if real a must have function or not
            pass

        self.excel_ini.status_update_src()
        #  add the program exit for the instrument control
        self.excel_ini.check_program_exit()

        pass
