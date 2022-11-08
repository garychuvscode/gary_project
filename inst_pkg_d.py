from os import supports_bytes_environ
import pyvisa
# # also for the jump out window, same group with win32con
import win32api
from win32con import MB_SYSTEMMODAL
# for float operation of string
import locale as lo
# for the delay function
import time
wait_samll = 0.05
wait_time = 0.2
# 100ms of waiting time for the separation of GPIB command
# maybe the instrument need the delay time
rm = pyvisa.ResourceManager()


class LPS_505N:
    # comments for the explanation of the power supply
    '''
    power supply
    ========
    vset0 => initial V;
    iset => initial I;
    ach_ch0 => active channel; \n
    GP_addr0 => GPIB address; \n
    state => initial state;
    '''

    # initialization is just for default parameter input, may not be able to use sub program
    # need to open and change output at other method in this object

    def __init__(self, vset0, iset0, act_ch0, GP_addr0, state0):
        # send value in when define the object
        self.vset_ini = vset0
        self.iset_ini = iset0
        self.act_ch_ini = act_ch0
        self.GP_addr_ini = GP_addr0
        self.state_ini = state0
        # add the initialization of none assign variable as 0, for simulation mode operation
        self.vset_o = 0
        self.iset_o = 0
        self.act_ch_o = 0
        self.state_o = 0
        self.cmd_str_out_sw = 0
        self.cmd_str_out_mode = 0
        self.cmd_str_V = 0
        self.cmd_str_I = 0
        self.iout_o = 0
        self.cmd_str_iout = 0
        # time.sleep(wait_samll)
        self.cmd_str_name = 0
        self.in_name = 0

        # use for ov and oc
        self.sup_ovp = 0
        self.sup_ocp = 0
        self.cmd_str_ovp = ''
        self.cmd_str_ocp = ''

        self.sim_inst = 1
        # simulation mode for the instrument
        # default set to high, in real mode, for the simulation mode,
        # change the control varable to 0
        # this will be put in each instrument object independently
        # and you will be able to switch to simulation mode any time you want

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the power supply, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'power supply simulation mode object'
            pass

    # should be better to use function directly for setting V and I

    # change output, one function is able to cover all the change

    def chg_out(self, vset1, iset1, act_ch1, state1):
        # if using the input value directly, no need to define the new name
        # for the input parameter
        # # but maybe still need, because access previous settings
        self.vset_o = vset1
        self.iset_o = iset1
        self.act_ch_o = act_ch1
        self.state_o = state1
        # using control to recover to off when finished command
        self.cmd_str_out_sw = "OUT" + \
            str(int(self.act_ch_o)) + ":" + str(self.state_o)
        # using normal(independent mode, not serial)
        self.cmd_str_out_mode = "OUT: normal"
        # setting string for the V and I
        self.cmd_str_V = ("PROG:VSET" + str(int(self.act_ch_o)) +
                          ":" + str(self.vset_o))
        self.cmd_str_I = ("PROG:ISET" + str(int(self.act_ch_o)) +
                          ":" + str(self.iset_o))
        # all command string using self because you can reference after the function is over,
        # it will left in the object, not disappear

        # write the command string for change power supply output

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.inst_obj.write(self.cmd_str_out_mode)
            self.inst_obj.write(self.cmd_str_V)
            self.inst_obj.write(self.cmd_str_I)
            self.inst_obj.write(self.cmd_str_out_sw)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('change output now with below GPIB string')
            print(str(self.cmd_str_out_mode))
            print(str(self.cmd_str_V))
            print(str(self.cmd_str_I))
            print(str(self.cmd_str_out_sw))

            pass

    # this function used to fast change voltage setting

    def change_V(self, vset2, channel_v):
        self.vset_o = vset2
        self.act_ch_o = channel_v
        # 20220127, add the channel index for the change_V function
        self.cmd_str_V = ("PROG:VSET" + str(int(self.act_ch_o)) +
                          ":" + str(self.vset_o))

        if self.sim_inst == 1:
            # must change both voltage and current together for every update
            # so the power supply will change output
            self.inst_obj.write(self.cmd_str_V)
            self.inst_obj.write(self.cmd_str_I)
            # when the source is already on, need to have the turn on command to update the final
            # setting to output, to prevent wrong behavior of the power supply
            # (only change the output voltage but not the current)
            # use another command to update at the same time
            self.inst_obj.write(self.cmd_str_out_sw)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('change V now with below GPIB string')
            print(str(self.cmd_str_V))
            print(str(self.cmd_str_I))
            print(str(self.cmd_str_out_sw))

            pass

        time.sleep(wait_samll)

    # this function used to fast change current setting

    def change_I(self, iset2, channel_i):
        self.iset_o = iset2
        self.act_ch_o = channel_i
        self.cmd_str_I = ("PROG:ISET" + str(int(self.act_ch_o)) +
                          ":" + str(self.iset_o))

        if self.sim_inst == 1:
            # must change both voltage and current together for every update
            # so the power supply will change output
            self.cmd_str_V = ("PROG:VSET" + str(int(self.act_ch_o)) +
                              ":" + str(self.vset_o))
            self.inst_obj.write(self.cmd_str_V)
            self.inst_obj.write(self.cmd_str_I)
            # when the source is already on, need to have the turn on command to update the final
            # setting to output, to prevent wrong behavior of the power supply
            # (only change the output voltage but not the current)
            # use another command to update at the same time
            self.cmd_str_out_sw = "OUT" + \
                str(int(self.act_ch_o)) + ":" + str(self.state_o)
            self.inst_obj.write(self.cmd_str_out_sw)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('change I now with below GPIB string')
            print(str(self.cmd_str_V))
            print(str(self.cmd_str_I))
            print(str(self.cmd_str_out_sw))

            pass

    # this function only used to check the simulation mode ouptput

    def read_iout(self, channel_i):
        self.act_ch_o = channel_i
        # self.iout_o = 0
        # not to initialize variable outside of the constructor
        # this will cause the erase of previous result
        self.cmd_str_iout = 'IOUT' + str(int(self.act_ch_o)) + '?'

        if self.sim_inst == 1:
            self.iout_o = self.inst_obj.query(self.cmd_str_iout)
            time.sleep(wait_samll)
            # after reading the iout from source, remove the A in the string
            self.iout_o = self.iout_o.replace('A', '')

            pass
        else:
            # for the simulatiom mode of change output
            print('change output now with below GPIB string')
            print(str(self.cmd_str_iout))
            print('now will retrun the sim current')
            self.iout_o = self.iout_o + 1
            print(self.iout_o)

            pass

        return str(self.iout_o)

    def vin_clibrate_singal_met(self, vin_ch, vin_target, met_v0, mcu0, excel0):

        # # stuff support coding
        # import inst_pkg_d as inst
        # import parameter_load_obj as par
        # import mcu_obj as mcu_main
        # mcu0 = mcu_main.MCU_control(0, 4)
        # # initial the object and set to simulation mode
        # met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
        # met_v0.sim_inst = 0
        # excel0 = par.excel_parameter('obj_main')

        # make sure all the type is correct
        vin_target = float(vin_target)
        vin_ch = int(vin_ch)

        # need to return the channel after the calibration is finished
        temp_mcu_channel = mcu0.meter_ch_ctrl
        v_res_temp = 0
        vin_diff = 0
        v_res_temp_f = 0
        # 20220103 add the Vin adjustment function
        # change the measure channel to the Vin channel (No.3)
        # this part added after the load turned on for better Vin setting

        mcu0.relay_ctrl(vin_ch)

        v_res_temp = met_v0.mea_v()
        v_res_temp_f = lo.atof(v_res_temp)
        # enter the calibration with real mode, otherwise pass the test and
        # go to the result directly
        if self.sim_inst == 1:

            # # measure the first Vin after relay change
            # if vin_ch == 0:
            #     v_res_temp = met_v0.mea_v()
            #     v_res_temp_f = lo.atof(v_res_temp)
            #     # the Vin calibration starts from here
            #     pass

            # elif vin_ch == 6:
            #     v_res_pwr_ch1 = met_v0.mea_v()
            #     v_res_temp_f = lo.atof(v_res_pwr_ch1)
            #     # the Vin calibration starts from here
            #     pass

            # elif vin_ch == 7:
            #     v_res_pwr_ch2 = met_v0.mea_v()
            #     v_res_temp_f = lo.atof(v_res_pwr_ch2)
            #     # the Vin calibration starts from here
            #     pass

            vin_diff = vin_target - v_res_temp_f
            vin_new = vin_target
            while vin_diff > excel0.vin_diff_set or vin_diff < (-1 * excel0.vin_diff_set):
                vin_new = vin_new + 0.5 * (vin_target - v_res_temp_f)
                # clamp for the Vin maximum
                if vin_new > excel0.pre_vin_max:
                    vin_new = excel0.pre_vin_max

                if vin_new < 0:
                    vin_new = 0

                if vin_ch == 0:
                    self.change_V(vin_new, excel0.relay0_ch)
                    # send the new Vin command for the auto testing channel
                    pass

                elif vin_ch == 6:
                    self.change_V(vin_new, excel0.relay6_ch)
                    # change the vsetting of channel 1 (mapped in program)
                    pass

                elif vin_ch == 7:
                    self.change_V(vin_new, excel0.relay7_ch)
                    # change the vsetting of channel 2 (mapped in program)
                    pass

                time.sleep(wait_time)
                # # measure the Vin change result
                # v_res_temp = met_v0.mea_v()
                # v_res_temp_f = lo.atof(v_res_temp)

                # 220524: update the calibration result to status variable

                v_res_temp = met_v0.mea_v()
                v_res_temp_f = lo.atof(v_res_temp)

                # if vin_ch == 0:
                #     v_res_temp = met_v0.mea_v()
                #     v_res_temp_f = lo.atof(v_res_temp)
                #     # the Vin calibration starts from here
                #     pass

                # elif vin_ch == 6:
                #     v_res_pwr_ch1 = met_v0.mea_v()
                #     v_res_temp_f = lo.atof(v_res_pwr_ch1)
                #     # the Vin calibration starts from here
                #     pass

                # elif vin_ch == 7:
                #     v_res_pwr_ch2 = met_v0.mea_v()
                #     v_res_temp_f = lo.atof(v_res_pwr_ch2)
                #     # the Vin calibration starts from here
                #     pass

                vin_diff = vin_target - v_res_temp_f

            # after the loop is finished, record the Vin meausred result (for full load) at the end
            # excel0.sh_org_tab.range((10, 9)).value = lo.atof(v_res_temp)
            # the Vin calibration ends from here

            # after the Vin calibration is finished, change the measure channel back
            # mcu0.meter_ch_ctrl = temp_mcu_channel
            # mcu0.relay_ctrl(mcu0.meter_ch_ctrl)
            mcu0.relay_ctrl(temp_mcu_channel)
            time.sleep(wait_time)
            # finished getting back to the initial state
        else:
            v_res_temp = lo.atof(v_res_temp)
            v_res_temp = v_res_temp + 1
            print('vin calibration sim mode, ' + str(v_res_temp) +' round ')
        # need to return the channel after the calibration is finished

        # the last measured value can also find in the meter result
        return str(v_res_temp)

    def sim_mode_out(self):
        print(self.vset_ini)
        print(self.iset_ini)
        print(self.act_ch_ini)
        print(self.state_ini)
        print(self.GP_addr_ini)
        print('')
        print(self.vset_o)
        print(self.iset_o)
        print(self.act_ch_o)
        print(self.state_o)
        print(self.iout_o)
        print('')
        print(self.cmd_str_out_sw)
        print(self.cmd_str_out_mode)
        print(self.cmd_str_I)
        print(self.cmd_str_V)
        print(self.cmd_str_iout)

    # consider to add the read V or read I? => but this may only for debugging

    def inst_close(self):
        print('to turn off all output and close GPIB device')

        if self.sim_inst == 1:
            # change all the output to 0V and 0A, for channel 1 to 3
            self.chg_out(0, 0, 1, 'off')
            self.chg_out(0, 0, 2, 'off')
            self.chg_out(0, 0, 3, 'off')
            # GPIB device close
            time.sleep(wait_samll)
            self.inst_obj.close()

            pass
        else:
            # for the simulatiom mode of change output
            print('all the channel turn off now')

            pass

    def inst_single_close(self, off_ch):
        print('to turn off related channel and "not" close GPIB device')

        if self.sim_inst == 1:
            # change all the output to 0V and 0A, for channel 1 to 3
            self.chg_out(0, 0, int(off_ch), 'off')
            print('turn off single channel of the power supply: CH_' + str(off_ch))
            # GPIB device close
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('simulation mode')
            print('turn off single channel of the power supply: CH_' + str(off_ch))

            pass

    def inst_name(self):
        # get the insturment name
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst_obj.query(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'pwr is in sim mode'

            pass

        return self.in_name

    # LPS 505 have the OVP and OCP can be set for incase
    def ov_oc_set(self, ovp0, ocp0):
        # parameter update for the ov and oc setting
        self.sup_ovp = ovp0
        self.sup_ocp = ocp0
        self.cmd_str_ovp = 'PORT:OVP:LEV ' + str(self.sup_ovp)
        self.cmd_str_ocp = 'PROT:OCP:LEV ' + str(self.sup_ocp)

        # command given to the instrument
        print(self.sup_ovp)
        print(self.sup_ocp)

        if self.sim_inst == 1:
            self.inst_obj.write(self.cmd_str_ovp)
            self.inst_obj.write(self.cmd_str_ocp)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('change OCP and OVP in sim mode')
            print(str(self.cmd_str_ovp))
            print(str(self.cmd_str_ocp))

            pass


class Met_34460:
    def __init__(self, mea_v_res0, max_mea_v0, mea_i_res0, max_mea_i0, GP_addr0):
        # assign the variable into object for saving
        self.mea_v_res_ini = mea_v_res0
        self.max_mea_v_ini = max_mea_v0
        self.mea_i_res_ini = mea_i_res0
        self.max_mea_i_ini = max_mea_i0
        self.GP_addr_ini = GP_addr0
        # better to define all the parameter at the initialization
        # to prevent error at runt time or the simulation mode
        self.cmd_str_mea_v = 0
        self.cmd_str_mea_i = 0
        self.mea_v_out = 0
        self.mea_i_out = 0
        self.cmd_str_name = 0

        self.sim_inst = 1
        # simulation mode for the instrument
        # default set to high, in real mode, for the simulation mode,
        # change the control varable to 0
        # this will be put in each instrument object independently
        # and you will be able to switch to simulation mode any time you want

    def open_inst(self):
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the meter, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'meter simulation mode object'
            pass

    def mea_v(self):
        # definie the command string and send the command string out to GPIB
        self.cmd_str_mea_v = (
            'MEASure:VOLT:DC? ' + str(self.max_mea_v_ini) + ',' + str(self.mea_v_res_ini))
        # save the result to below variable
        # self.mea_v_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_v_out = self.mea_v_out + 1

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.mea_v_out = self.inst_obj.query(self.cmd_str_mea_v)
            print('v measure is: ' + self.mea_v_out)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('measure votlage with below GPIB string')
            print(str(self.cmd_str_mea_v))
            self.mea_v_out = self.mea_v_out + 1
            print(self.mea_v_out)

            pass

        # return the result back to main program, should be able to access again in this object
        return str(self.mea_v_out)

    def mea_v2(self, v_max=0, v_level=0):
        # 220520: if input 0, it will be same with the original version
        # definie the command string and send the command string out to GPIB
        if v_max == 0:
            # if there are no vmax setting input, set to default when meter is defined
            v_max = self.max_mea_v_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass
        if v_level == 0:
            # if there are no vmax setting input, set to default when meter is defined
            v_level = self.mea_v_res_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass

        self.cmd_str_mea_v = (
            'MEASure:VOLT:DC? ' + str(v_max) + ',' + str(v_level))
        # save the result to below variable
        # self.mea_v_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_v_out = self.mea_v_out + 1
        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.mea_v_out = self.inst_obj.query(self.cmd_str_mea_v)
            print('V measure V2 is: ' + self.mea_v_out)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('measure votlage with below GPIB string (mea_V2)')
            print(str(self.cmd_str_mea_v))
            self.mea_v_out = self.mea_v_out + 1
            print(self.mea_v_out)

            pass
        # return the result back to main program, should be able to access again in this object
        return str(self.mea_v_out)

    def mea_i(self):
        # define command string
        self.cmd_str_mea_i = (
            'MEASure:CURR:DC? ' + str(self.max_mea_i_ini) + ',' + str(self.mea_i_res_ini))
        # save result to below
        # self.mea_i_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_i_out = self.mea_i_out + 1

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.mea_i_out = self.inst_obj.query(self.cmd_str_mea_i)
            print('i measure is: ' + self.mea_i_out)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('measure current with below GPIB string')
            print(str(self.cmd_str_mea_i))
            self.mea_i_out = self.mea_i_out + 1
            print(self.mea_i_out)

            pass

        # return back and save in object
        return str(self.mea_i_out)

    def mea_i2(self, i_max=0, i_level=0):

        if i_max == 0:
            # if there are no vmax setting input, set to default when meter is defined
            i_max = self.max_mea_i_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass
        if i_level == 0:
            # if there are no vmax setting input, set to default when meter is defined
            i_level = self.mea_i_res_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass

        # define command string
        self.cmd_str_mea_i = (
            'MEASure:CURR:DC? ' + str(i_max) + ',' + str(i_level))
        # save result to below
        # self.mea_i_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_i_out = self.mea_i_out + 1
        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.mea_i_out = self.inst_obj.query(self.cmd_str_mea_i)
            print('I measure V2 is: ' + self.mea_i_out)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('measure current with below GPIB string (mea_I2)')
            print(str(self.cmd_str_mea_i))
            self.mea_i_out = self.mea_i_out + 1
            print(self.mea_i_out)

            pass
        # return back and save in object
        return str(self.mea_i_out)

    def sim_mode_out(self):
        # simulation output mode to check result in terminal
        print(self.mea_v_res_ini)
        print(self.max_mea_v_ini)
        print(self.mea_i_res_ini)
        print(self.max_mea_i_ini)
        print(self.GP_addr_ini)
        print('')
        print(self.inst_obj)
        print(self.cmd_str_mea_v)
        print(self.cmd_str_mea_i)
        print('')
        print(self.mea_v_out)
        print(self.mea_i_out)

    def inst_close(self):
        print('to turn off all output and close GPIB device')

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            time.sleep(wait_samll)
            # change all the output to 0V and 0A, for channel 1 to 3
            self.inst_obj.close()

            pass
        else:
            # for the simulatiom mode of change output
            print('close the meter ')

            pass

    def inst_name(self):
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst_obj.query(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'met is in sim mode'

            pass
        return self.in_name


class chroma_63600:
    # initialization is just for default parameter input, may not be able to use sub program
    # need to open and change output at other method in this object

    def __init__(self, act_ch0, GP_addr0, mode0):
        # send value in when define the object
        self.act_ch_ini = act_ch0
        self.GP_addr_ini = GP_addr0
        # CCH CCL or other selection (default mode all the same)
        self.mode_ini = [mode0, mode0, mode0, mode0]
        # add the initialization of none assign variable as 0, for simulation mode operation
        self.act_ch_o = 1
        # array of current setting, map number = number -1
        self.i_sel_ch = [0, 0, 0, 0]
        # current calibration index for each channel => number -1
        self.i_cal_offset_ch = [0, 0, 0, 0]
        # current calibration index for each channel => number -1
        self.i_cal_leakage_ch = [0, 0, 0, 0]
        # self.i_sel_ch1 = 0
        # self.i_sel_ch2 = 0
        # self.i_sel_ch3 = 0
        # self.i_sel_ch4 = 0
        self.mode_o = self.mode_ini
        # only change mode o, not change mode_ini, data loaded from mode o

        # calibration mode setting for loader
        # default turn off the calibration mode
        self.cal_mode_en = 0

        self.state_o = ['off', 'off', 'off', 'off']
        # status also use array to define
        # need to add space => 'Load on' or 'load off'
        self.cmd_str_ch_set = 0
        self.cmd_str_status = 0
        self.cmd_str_I_load = 0
        self.cmd_str_I_read = 0
        self.cmd_str_V_read = 0
        self.cmd_str_mode_set = 0

        # time.sleep(wait_samll)
        # need to keep this for overy instrument, it's for reading the name of instrument
        self.cmd_str_name = ''
        self.in_name = ''

        # other parameter needed in loader class
        self.v_out = 0
        self.i_out = 0

        # loader object definition
        self.inst_obj = 0
        # other definition
        self.errflag = 0
        # error flag indicate => 0 is ok and 1 is error

        self.sim_inst = 1
        # simulation mode for the instrument
        # default set to high, in real mode, for the simulation mode,
        # change the control varable to 0
        # this will be put in each instrument object independently
        # and you will be able to switch to simulation mode any time you want

    # need to watch out if the power limit, current limit need to be set and config through
    # the different mode of the load setting, to prevent crash of the load during auto testing
    # the accuracy and current operating range will also be different for the different mode

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        # here is to define object map to the loader
        # conntect to the related GPIB address
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the meter, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'loader simulation mode object'
            pass

        pass

    # should be better to use function directly for setting V and I

    # change output, one function is able to cover all the change

    def chg_out(self, iset1, act_ch1, state1):
        # if using the input value directly, no need to define the new name
        # for the input parameter
        # # but maybe still need, because access previous settings
        self.act_ch_o = act_ch1
        # update the current and state setting in related channel for record in the class
        # record in class can prevent error change or setting when not refresh
        # only change one in each sub program call, adjust one channel at a time

        # need to make sure all the input become correct type
        iset1 = float(iset1)
        act_ch1 = int(act_ch1)

        # filter iset before update current into the array, based on the CCx mode
        # CCL => 200mA, CCM => 2A, CCH => 20A
        # to limit the iset at selected range and is able to find out wrong
        if self.mode_o[int(self.act_ch_o) - 1] == 'CCH':
            if iset1 > 20:
                iset1 = 20
                self.errflag = 1
        elif self.mode_o[int(self.act_ch_o) - 1] == 'CCM':
            if iset1 > 2:
                iset1 = 2
                self.errflag = 1
        elif self.mode_o[int(self.act_ch_o) - 1] == 'CCL':
            if iset1 > 0.2:
                iset1 = 0.2
                self.errflag = 1

        self.i_sel_ch[int(self.act_ch_o) - 1] = iset1
        print(self.i_sel_ch)
        self.state_o[int(self.act_ch_o) - 1] = state1
        print(self.state_o)
        # decide witch channel to use first
        self.cmd_str_ch_set = "CHAN " + str(int(self.act_ch_o))
        print(self.cmd_str_ch_set)
        # mode setting update map with the array
        self.cmd_str_mode_set = "MODE " + \
            str(self.mode_o[int(self.act_ch_o) - 1])
        print(self.cmd_str_mode_set)
        # update the string for current definition (map from the related array element)

        # 220921: reduce the error cause from leakage
        self.cmd_str_I_load = "curr:stat:L1 " + \
            str(self.i_sel_ch[int(self.act_ch_o) - 1])
        print(self.cmd_str_I_load)

        if self.cal_mode_en == 1:
            self.cmd_str_I_load = "curr:stat:L1 " + \
                str(self.i_sel_ch[int(self.act_ch_o) - 1] -
                    self.i_cal_leakage_ch[int(self.act_ch_o) - 1])
            print('after calibration')
            print(self.cmd_str_I_load)

        # update the status of on and off
        self.cmd_str_status = ("Load " + self.state_o[int(self.act_ch_o) - 1])
        print(self.cmd_str_status)
        # all command string using self because you can reference after the function is over,
        # it will left in the object, not disappear

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            # write the command string for change power supply output
            # writring sequence: channel => current setting => status update
            # check if it changes like power supply? need to update status to refresh command
            self.inst_obj.write(self.cmd_str_ch_set)
            self.inst_obj.write(self.cmd_str_mode_set)
            self.inst_obj.write(self.cmd_str_I_load)
            self.inst_obj.write(self.cmd_str_status)
            # add the break point here to double if the command update based on the write command of status
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('change loader output now with below GPIB string')
            print(str(self.cmd_str_ch_set))
            print(str(self.cmd_str_mode_set))
            print(str(self.cmd_str_I_load))
            print(str(self.cmd_str_status))

            pass

        return self.errflag
        # return error when there are load current setting error
        # it should also be ok to not setting the return variable

    def chg_out2(self, iset1, act_ch1, state1):
        '''
        202211 new version, auto change CCH, CCM and CCL based on iset
        '''
        self.act_ch_o = act_ch1
        # update the current and state setting in related channel for record in the class
        # record in class can prevent error change or setting when not refresh
        # only change one in each sub program call, adjust one channel at a time

        # need to make sure all the input become correct type
        iset1 = float(iset1)
        act_ch1 = int(act_ch1)

        # CCL => 200mA, CCM => 2A, CCH => 20A
        # auto change iset based on different loading requirement
        if iset1 > 2:
            # this should be CCM mode
            self.chg_mode(act_ch1, "CCH")
            if iset1 > 20:
                iset1 = 20
        elif iset1 > 0.2:
            # this should be CCM mode
            self.chg_mode(act_ch1, "CCM")
        else:
            # this should be CCL mode
            self.chg_mode(act_ch1, "CCL")

        self.i_sel_ch[int(self.act_ch_o) - 1] = iset1
        print(self.i_sel_ch)
        self.state_o[int(self.act_ch_o) - 1] = state1
        print(self.state_o)
        # decide witch channel to use first
        self.cmd_str_ch_set = "CHAN " + str(int(self.act_ch_o))
        print(self.cmd_str_ch_set)
        # mode setting update map with the array
        self.cmd_str_mode_set = "MODE " + \
            str(self.mode_o[int(self.act_ch_o) - 1])
        print(self.cmd_str_mode_set)
        # update the string for current definition (map from the related array element)

        # 220921: reduce the error cause from leakage
        self.cmd_str_I_load = "curr:stat:L1 " + \
            str(self.i_sel_ch[int(self.act_ch_o) - 1])
        print(self.cmd_str_I_load)

        if self.cal_mode_en == 1:
            self.cmd_str_I_load = "curr:stat:L1 " + \
                str(self.i_sel_ch[int(self.act_ch_o) - 1] -
                    self.i_cal_leakage_ch[int(self.act_ch_o) - 1])
            print('after calibration')
            print(self.cmd_str_I_load)

        # update the status of on and off
        self.cmd_str_status = ("Load " + self.state_o[int(self.act_ch_o) - 1])
        print(self.cmd_str_status)
        # all command string using self because you can reference after the function is over,
        # it will left in the object, not disappear

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            # write the command string for change power supply output
            # writring sequence: channel => current setting => status update
            # check if it changes like power supply? need to update status to refresh command
            self.inst_obj.write(self.cmd_str_ch_set)
            self.inst_obj.write(self.cmd_str_mode_set)
            self.inst_obj.write(self.cmd_str_I_load)
            self.inst_obj.write(self.cmd_str_status)
            # add the break point here to double if the command update based on the write command of status
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('change loader output now with below GPIB string')
            print(str(self.cmd_str_ch_set))
            print(str(self.cmd_str_mode_set))
            print(str(self.cmd_str_I_load))
            print(str(self.cmd_str_status))

            pass

        return self.errflag
        # return error when there are load current setting error
        # it should also be ok to not setting the return variable

    # this function used to read the feedback voltage measurement

    def read_vout(self, act_ch1):
        # also need to input the channel to check, and it will update the final channel
        self.act_ch_o = act_ch1

        # string setting and update
        self.cmd_str_ch_set = "CHAN " + str(int(self.act_ch_o))
        self.cmd_str_V_read = ("MEAS:VOLT?")

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            # string write and action
            self.inst_obj.write(self.cmd_str_ch_set)
            self.v_out = self.inst_obj.query(self.cmd_str_V_read)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('loader read the voltage with below GPIB string')
            print(str(self.cmd_str_ch_set))
            print(str(self.cmd_str_V_read))
            self.v_out = 'sim_vout'
            # to show the program had been run through this place
            # the result change after run through

            pass

        return str(self.v_out)

    # this function used to feedback the output current measurement

    def read_iout(self, act_ch1):
        # also need to input the channel to check, and it will update the final channel
        self.act_ch_o = act_ch1

        # string setting and update
        self.cmd_str_ch_set = "CHAN " + str(int(self.act_ch_o))
        self.cmd_str_V_read = ("MEAS:CURR?")

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            # string write and action
            self.inst_obj.write(self.cmd_str_ch_set)
            self.i_out = self.inst_obj.query(self.cmd_str_V_read)
            time.sleep(wait_samll)

            if self.cal_mode_en == 1:
                # 1 is constant calibration mode

                temp_i_out = lo.atof(self.i_out)
                print('in cal mode of loader => mode 1 ')
                print('before:')
                print(temp_i_out)

                temp_i_out = temp_i_out - \
                    self.i_cal_offset_ch[self.act_ch_o - 1]
                print('after:')
                print(temp_i_out)

                # and current can't be 0
                if temp_i_out < 0:
                    temp_i_out = 0
                    pass

                self.i_out = str(temp_i_out)
                print(self.i_out)

            pass
        else:
            # for the simulatiom mode of change output
            print('loader read current with below GPIB string')
            print(str(self.cmd_str_ch_set))
            print(str(self.cmd_str_V_read))
            self.i_out = self.i_out + 1

            pass

        # 220330: chroma response don't have A in the string, only the power supply have string
        # # after reading the iout from source, remove the A in the string
        # self.iout_o = self.iout_o.replace('A', '')
        return str(self.i_out)

    def chg_out_auto_mode(self, iset1, act_ch1, state1):

        if iset1 < 0.2:
            self.chg_mode(act_ch1, 'CCL')
            pass
        elif iset1 > 0.2 and iset1 < 2:
            self.chg_mode(act_ch1, 'CCM')
            # change mode will turn off first
            # need to turn on again after change the mode
            pass
        elif iset1 > 2 and iset1 < 20:
            self.chg_mode(act_ch1, 'CCH')
            pass

        self.chg_out(iset1, act_ch1, state1)
        # after changing the mode, any current can be accept for the loader setting
        # just mode will change
        # refer to the iset1

        pass

    def sim_mode_out(self):
        print(self.act_ch_ini)
        print(self.GP_addr_ini)
        print(self.mode_ini)
        print('')
        print(self.act_ch_o)
        print(self.i_sel_ch)
        print(self.mode_o)
        print(self.state_o)
        print('')
        print(self.cmd_str_ch_set)
        print(self.cmd_str_status)
        print(self.cmd_str_I_load)
        print(self.cmd_str_I_read)
        print(self.cmd_str_V_read)
        print('')
        print(self.cmd_str_name)
        print(self.in_name)
        print(self.v_out)
        print(self.i_out)
        print('')
        print(self.inst_obj)
        print(self.errflag)
        pass

    def chg_state_single(self, channel0, state0):
        self.act_ch_o = int(channel0)
        self.state_o[self.act_ch_o - 1] = state0
        self.chg_out(self.i_sel_ch[self.act_ch_o - 1],
                     self.act_ch_o, self.state_o[self.act_ch_o - 1])

        pass

    def chg_curr_single(self, channel0, curr0):
        self.act_ch_o = channel0
        self.i_sel_ch[self.act_ch_o - 1] = curr0
        self.chg_out(self.i_sel_ch[self.act_ch_o - 1],
                     self.act_ch_o, self.state_o[self.act_ch_o - 1])

        pass

    # consider to add the read V or read I? => but this may only for debugging

    # to close all the channel and turn off the GPIB connection, need to open again

    def inst_close(self):
        print('to turn off all output and close GPIB device')

        if self.sim_inst == 1:
            # turn off the 4 channel of loader
            self.chg_out(0, 1, 'off')
            self.chg_out(0, 2, 'off')
            self.chg_out(0, 3, 'off')
            self.chg_out(0, 4, 'off')
            # GPIB device close
            time.sleep(wait_samll)
            self.inst_obj.close()

            pass
        else:
            # for the simulatiom mode of change output
            print('all the loader channel turn off now')

            pass

    def inst_single_close(self, off_ch):
        print('to turn off related channel and "not" close GPIB device')

        if self.sim_inst == 1:
            # change all the output to 0V and 0A, for channel 1 to 3
            self.chg_out(0, int(off_ch), 'off')
            print('turn off single channel of the loader: CH_' + str(off_ch))
            # GPIB device close
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('simulation mode')
            print('turn off single channel of the power supply: CH_' + str(off_ch))

            pass

    # used to get the instrument name and return to main program

    def inst_name(self):
        # get the insturment name
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst_obj.query(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'loader is in sim mode'

            pass

        return self.in_name

    # used to change output mode

    def chg_mode(self, act_ch1, mode0):
        self.act_ch_o = act_ch1

        # lock the mode selection only in CC mode, set to CCH if there are error
        if mode0 == 'CCL' or mode0 == 'CCM' or mode0 == 'CCH':
            self.mode_o[int(act_ch1)-1] = mode0

        else:
            mode0 = 'CCH'
            self.mode_o[int(act_ch1)-1] = mode0
            self.errflag = 1

        # need to turn off channel of there are mode change
        self.chg_out(0, self.act_ch_o, 'off')

        # !! dont have other mode setting cmd string in the chg out, it should be based on the mode to choose different command string
        # lock down mode selection for only CC mode in this object
        return self.errflag

    def error_clear(self):
        self.errflag = 0
        # this sub is used to clear error flag
        pass

    def current_calibration(self, met_v0, pwr0, pwr_ch, load_channel, v_test):
        # current calubration function for loader
        # meed meter(met0) amd power supply(pwr0)
        # object input for control

        # # stuff support coding
        # import inst_pkg_d as inst
        # pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
        # pwr0.sim_inst = 0
        # # initial the object and set to simulation mode
        # met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
        # met_v0.sim_inst = 0

        # since there are offset and leakage at the chroma loader
        # get both setting updatte in this case

        # read from loader
        read_i_result = 0
        # meter result
        met_i_result = 0
        # result
        calibration_result_offset = 0
        calibration_result_leakage = 0
        # get 20 times for calibration
        average = 20

        # first to check setup
        # pwr - met_i - loader
        self.act_ch_o = load_channel
        # disable current calibration during doing calibration
        cal_mode_temp = self.cal_mode_en
        self.cal_mode_en = 0

        print('window jump out')
        self.msg_res = win32api.MessageBox(0, 'press enter if hardware configuration is correct',
                                           'pwr, met, loader ch_' + str(load_channel) + ' all in series for calibration current')

        # calibration start from here

        pwr0.chg_out(v_test, 0.5, pwr_ch, 'on')
        self.chg_out(0, load_channel, 'on')

        x_cal = 0
        # for the offset calculation
        result_temp = 0
        # for the leakage calculation
        met_temp = 0
        while x_cal < average:

            met_i_result = met_v0.mea_i()
            print(met_i_result)
            time.sleep(wait_samll)
            read_i_result = self.read_iout(load_channel)
            time.sleep(wait_samll)

            # measure I(loader) - calibration = measure I(meter)
            print('measure I(loader) - calibration = measure I(meter)')
            result_temp = result_temp + \
                lo.atof(read_i_result) - lo.atof(met_i_result)
            met_temp = met_temp + lo.atof(met_i_result)

            x_cal = x_cal + 1
            pass
        calibration_result_offset = result_temp / average
        calibration_result_leakage = met_temp / average

        self.i_cal_offset_ch[load_channel - 1] = calibration_result_offset
        self.i_cal_leakage_ch[load_channel - 1] = calibration_result_leakage
        print('calibration for ' + str(average) + ' times finished')
        print('to get along well with lover well, we need to calibrate offtenly')

        print('offset is: ')
        print(self.i_cal_offset_ch[load_channel - 1])
        print('leakage is: ')
        print(self.i_cal_leakage_ch[load_channel - 1])

        print('window jump out')
        self.msg_res = win32api.MessageBox(0, 'press enter for next step',
                                           'the offset and leakage of ch' + str(load_channel) + ' show in termainal')

        pwr0.chg_out(0, 0.5, pwr_ch, 'off')
        self.chg_out(0, load_channel, 'off')

        # return the mode setting when the calibration is finished
        self.cal_mode_en = cal_mode_temp

        pass

    def current_cal_setup(self, off_ch1, off_ch2, off_ch3, off_ch4, leak_ch1, leak_ch2, leak_ch3, leak_ch4):
        # this sub function used to setup the current calibartion constant directly
        # from the excel input
        self.i_cal_offset_ch[0] = off_ch1
        self.i_cal_offset_ch[1] = off_ch2
        self.i_cal_offset_ch[2] = off_ch3
        self.i_cal_offset_ch[3] = off_ch4

        self.i_cal_leakage_ch[0] = leak_ch1
        self.i_cal_leakage_ch[1] = leak_ch2
        self.i_cal_leakage_ch[2] = leak_ch3
        self.i_cal_leakage_ch[3] = leak_ch4

        pass


# still on planning


class Keth_2440:
    # initialization is just for default parameter input, may not be able to use sub program
    # need to open and change output at other method in this object

    def __init__(self, vset0, iset0, GP_addr0, state0, source_type, clamp_VI):
        # send value in when define the object
        self.vset_ini = vset0
        # vset is the V_clamp of source meter
        self.iset_ini = iset0
        self.GP_addr_ini = GP_addr0
        self.state_ini = state0
        self.source_type_ini = source_type
        self.source_type_inib = 0
        # source type is to choose V source or I source
        self.clamp_VI_ini = clamp_VI
        # clamp V is in V and clamp I is in A
        # add the initialization of none assign variable as 0, for simulation mode operation
        self.source_type_o = source_type
        if self.source_type_o == 'CURR':
            self.source_type_ob = 'VOLT'
        elif self.source_type_o == 'VOLT':
            self.source_type_ob = 'CURR'
        self.source_type_ob = 0
        self.vset_o = 0
        self.iset_o = 0
        self.state_o = 0
        # string for on or off
        # ---

        self.turn_off_str = 'OUTP OFF'
        self.turn_on_str = 'OUTP ON'
        self.read_str = 'READ?'
        self.read_v_str1 = "SENS:FUNC 'VOLT'"
        self.read_v_str2 = "SENS:VOLT:RANG:AUTO ON"
        self.read_v_str3 = 'FORM:ELEM VOLT'
        self.read_i_str1 = "SENS:FUNC 'CURR'"
        self.read_i_str2 = "SENS:CURR:RANG:AUTO ON"
        self.read_i_str3 = 'FORM:ELEM CURR'

        self.read_res = 0
        # read command for source meter
        # time.sleep(wait_samll)
        self.cmd_str_name = 0
        # instrument name
        self.in_name = 0
        self.clamp_VI_o = 0
        self.read_mode = self.source_type_ini

        self.sim_inst = 1
        # simulation mode for the instrument
        # default set to high, in real mode, for the simulation mode,
        # change the control varable to 0
        # this will be put in each instrument object independently
        # and you will be able to switch to simulation mode any time you want

        #  this is the result string for the result return of simulation mode
        self.return_str = 0

    # start from the source meter, to get a better efficiency when coding,
    # refresh the query and write to sub program(which include command print)
    # reduce the coding structure

    def query_write(self, cmd_str0):
        print(cmd_str0)

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.return_str = self.inst_obj.query(cmd_str0)

            pass
        else:
            # for the simulatiom mode of change output
            print('query write for sub program')
            self.return_str = self.return_str + 1

            pass

        return str(self.return_str)

    def only_write(self, cmd_str1):
        print(cmd_str1)

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.inst_obj.write(cmd_str1)

            pass
        else:
            # for the simulatiom mode of change output
            print('only write for sub program')

            pass

    def load_off(self):
        # turn off the load is independent command,
        # turn on the load is integrated in the change of command
        print('now is going to turn off')
        # self.inst_obj.write(self.turn_off_str)

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.inst_obj.write(self.turn_off_str)

            pass
        else:
            # for the simulatiom mode of change output
            print('source meter is going to turn off')

            pass

        # remember to change the state variable so change type won't have error
        self.state_o = 'off'

    def load_on(self):
        # turn off the load is independent command,
        # turn on the load is integrated in the change of command
        print('now is going to turn on')

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.inst_obj.write(self.turn_on_str)

            pass
        else:
            # for the simulatiom mode of change output
            print('source meter is going to turn on')

            pass

        # remember to change the state variable so change type won't have error
        self.state_o = 'on'

    def ini_inst(self):

        str_temp = ''
        # when need to re-run initialization from during operation
        # after change mode or other settings
        str_temp = 'SOUR:CLE:AUTO off'
        self.only_write(str_temp)
        str_temp = 'SOUR:FUNC ' + str(self.source_type_o)
        self.only_write(str_temp)
        str_temp = 'SOUR:' + str(self.source_type_o) + ':MODE FIXED'
        self.only_write(str_temp)
        str_temp = "SENS:FUNC '" + str(self.source_type_o) + "'"
        self.only_write(str_temp)
        str_temp = "SENS:" + str(self.source_type_o) + ":RANG:AUTO ON"
        self.only_write(str_temp)
        if self.source_type_o == 'CURR':
            self.source_type_ob = 'VOLT'
        elif self.source_type_o == 'VOLT':
            self.source_type_ob = 'CURR'
        str_temp = "FORM:ELEM " + str(self.source_type_o)
        self.only_write(str_temp)
        str_temp = "SENS:" + str(self.source_type_ob) + \
            ":PROT " + str(self.clamp_VI_o)
        self.only_write(str_temp)

    def open_inst(self):
        str_temp = ''
        # maybe no need to define rm for global variable
        # global rm
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the source meter, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'SRC meter simulation mode object'
            pass
        # when the instrument open, need to get some intitalization

        # config source meter as a current(CURR)/voltage(VOLT) source
        str_temp = '*RST'
        self.only_write(str_temp)
        str_temp = 'SOUR:CLE:AUTO off'
        self.only_write(str_temp)
        str_temp = 'SOUR:FUNC ' + str(self.source_type_ini)
        self.only_write(str_temp)
        str_temp = 'SOUR:' + str(self.source_type_ini) + ':MODE FIXED'
        self.only_write(str_temp)
        str_temp = "SENS:FUNC '" + str(self.source_type_ini) + "'"
        self.only_write(str_temp)
        str_temp = "SENS:" + str(self.source_type_ini) + ":RANG:AUTO ON"
        self.only_write(str_temp)
        str_temp = "FORM:ELEM " + str(self.source_type_ini)
        self.only_write(str_temp)
        if self.source_type_ini == 'CURR':
            self.source_type_inib = 'VOLT'
        elif self.source_type_ini == 'VOLT':
            self.source_type_inib = 'CURR'
        str_temp = "SENS:" + str(self.source_type_inib) + \
            ":PROT " + str(self.clamp_VI_ini)
        self.only_write(str_temp)
        # default off after initialization
        self.load_off()

    def change_type(self, type_temp, clamp_vi):
        # this function is used to change source type between I source or V source
        # type can only be change when output is already turned off
        if self.state_o == 'on':
            # 220521: if status is keep on, can only change clamp
            self.clamp_VI_o = clamp_vi
            self.ini_inst()
            pass
            # bypass the change_type request if state is on
        elif self.state_o == 'off':
            # change the type and clamp parameter mapping during off state
            # the next command send will be based on the state setting
            self.source_type_o = type_temp
            self.clamp_VI_o = clamp_vi
            self.ini_inst()
            # after change control parameter and re-run initialization, type change ok

    # this function used to operate as V source and change V

    def change_V(self, vset2, state2):
        if state2 == 'keep':
            # change the setting without changing status
            # if it's on, keep on, if it's off, keep off
            pass

        elif state2 == 'off':
            self.load_off()
            # change the clamp setting after turn off the output
            # watch out that re-turn on the load is needed if change with off
            print('a re-turn on of command is needed after off settings')

        if self.source_type_o == 'CURR':
            # if the source type is in current source
            # change the voltage clamp settings
            self.clamp_VI_o = vset2
            self.ini_inst()
            time.sleep(wait_samll)
        # reset the clamp voltage and reset run the initialization
        elif self.source_type_o == 'VOLT':
            self.vset_o = vset2
            str_temp = 'SOUR:' + \
                str(self.source_type_o) + ':RANG ' + str(vset2)
            self.only_write(str_temp)
            str_temp = 'SOUR:' + \
                str(self.source_type_o) + ':LEV ' + str(vset2)
            self.only_write(str_temp)
            time.sleep(wait_samll)

        if state2 == 'on':
            self.load_on()
            print('load is turn on after the setting is finished')

    # this function used to fast change current setting

    def change_I(self, iset2, state2):
        if state2 == 'keep':
            # change the setting without changing status
            # if it's on, keep on, if it's off, keep off
            pass

        elif state2 == 'off':
            self.load_off()
            # change the clamp setting after turn off the output
            # watch out that re-turn on the load is needed if change with off
            print('a re-turn on of command is needed after off settings')

        if self.source_type_o == 'VOLT':
            # if the source type is in current source
            # change the voltage clamp settings
            self.clamp_VI_o = iset2
            self.ini_inst()
            time.sleep(wait_samll)
        # reset the clamp voltage and reset run the initialization
        elif self.source_type_o == 'CURR':
            self.iset_o = iset2
            str_temp = 'SOUR:' + \
                str(self.source_type_o) + ':RANG ' + str(iset2)
            self.only_write(str_temp)
            str_temp = 'SOUR:' + \
                str(self.source_type_o) + ':LEV ' + str(iset2)
            self.only_write(str_temp)
            time.sleep(wait_samll)

        if state2 == 'on':
            self.load_on()
            print('load is turn on after the setting is finished')

    # this function only used to check the simulation mode ouptput

    def read(self, read_type):
        if self.state_o == 'on':
            # choose to read voltage or current
            if read_type == 'VOLT':
                self.only_write(self.read_v_str1)
                self.only_write(self.read_v_str2)
                self.only_write(self.read_v_str3)
                self.read_mode = read_type
            elif read_type == "CURR":
                self.only_write(self.read_i_str1)
                self.only_write(self.read_i_str2)
                self.only_write(self.read_i_str3)
                self.read_mode = read_type

            # self.read_res = self.inst_obj.query(self.read_str)
            self.read_res = self.query_write(self.read_str)
            time.sleep(wait_samll)
            # after reading the iout from source, remove the A in the string
            if self.sim_inst == 1:
                self.read_res = self.read_res.replace('A', '')
                pass
            # # 220921 add the read current exception
            # # source meter can't used this rule, positive and negative current
            # if read_type == "CURR" :
            #     if lo.atof(self.read_res) < 0 :
            #         self.read_res = 0
            #         pass
            #     pass

            print('mode ' + str(self.source_type_o) +
                  ', the ' + str(self.read_mode) + ' reading result is: ' + str(self.read_res))
            return str(self.read_res)
        else:
            return '0'
            pass
            # since source meter is unable to read during off state
            # need to pass the read function when meter is off

    def sim_mode_out(self):
        print('vset_ini is:' + str(self.vset_ini))
        print('iset_ini is:' + str(self.iset_ini))
        print('source_type_ini is:' + str(self.source_type_ini))
        print('source_type_inib is:' + str(self.source_type_inib))
        print('state_ini is:' + str(self.state_ini))
        print('GP_addr_ini is:' + str(self.GP_addr_ini))
        print('clamp_VI_ini is:' + str(self.clamp_VI_ini))
        print('')
        print('vset_o is:' + str(self.vset_o))
        print('iset_o is:' + str(self.iset_o))
        print('source_type_o is:' + str(self.source_type_o))
        print('source_type_ob is:' + str(self.source_type_ob))
        print('state_o is:' + str(self.state_o))
        print('clamp_VI_o is:' + str(self.clamp_VI_o))
        print('')
        print('turn_off_str is:' + str(self.turn_off_str))
        print('turn_on_str is:' + str(self.turn_on_str))
        print('read_str is:' + str(self.read_str))
        print('cmd_str_name is:' + str(self.cmd_str_name))
        print('in_name is:' + str(self.in_name))

    # consider to add the read V or read I? => but this may only for debugging

    def inst_close(self):
        print('to turn off all output and close GPIB device')
        self.load_off()

        if self.sim_inst == 1:
            # GPIB device close
            time.sleep(wait_samll)
            self.inst_obj.close()

            pass
        else:
            # for the simulatiom mode of change output
            print('source meter 2440 close now ')

            pass

    def inst_name(self):
        # get the insturment name
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst_obj.query(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'src is in sim mode'

            pass

        return self.in_name


class chamber_su242:
    # initialization is just for default parameter input, may not be able to use sub program
    # need to open and change output at other method in this object

    def __init__(self, tset0, GP_addr0, state0, l_limt_0, h_limt_0, ready_err_0):
        # send value in when define the object
        self.tset_ini = tset0
        # vset is the V_clamp of source meter
        self.GP_addr_ini = GP_addr0
        self.state_ini = state0
        self.temp_L_limt_ini = l_limt_0
        self.temp_H_limt_ini = h_limt_0
        self.ready_err_ini = ready_err_0

        self.tset_o = 0
        # the last set temperature
        self.tmea_o = 0
        # the last read temperature from chamber
        self.state_o = state0
        self.state_o_value = []
        # return type will be
        # measured, set, upper limit, lower limit
        self.temp_o = 0
        # current temperature
        self.temp_L_limt = l_limt_0
        self.temp_H_limt = h_limt_0
        self.ready_err = ready_err_0
        # string for on or off
        # ---

        self.mode_set_str = 'MODE, CONSTANT'
        self.temp_set_str = 'TEMP, S'
        self.end_str = '\n'
        self.temp_L_limt_set_str = "TEMP, L"
        self.temp_H_limt_set_str = "TEMP, H"
        self.temp_set_str_H = ' H'
        self.temp_set_str_L = " L"
        self.temp_read_str = "TEMP?"
        self.turn_off_str = 'MODE, STANDBY'

        self.sim_inst = 1
        # simulation mode for the instrument
        # default set to high, in real mode, for the simulation mode,
        # change the control varable to 0
        # this will be put in each instrument object independently
        # and you will be able to switch to simulation mode any time you want

        self.cmd_str_name = ''

    def query_write(self, cmd_str0):
        print(cmd_str0)

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            return_str = self.inst_obj.query(cmd_str0, 0.5)

            pass
        else:
            # for the simulatiom mode of change output
            print('query write for sub program')
            return_str = 'chamber_sim_mode return query'

            pass
        print(return_str)
        return return_str

    def only_write(self, cmd_str1):
        print(cmd_str1)

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.inst_obj.write(cmd_str1)

            pass
        else:
            # for the simulatiom mode of change output
            print('only write for sub program')

            pass

    def chamber_off(self):
        # turn off the load is independent command,
        # turn on the load is integrated in the change of command
        print('now is going to turn off')
        # self.inst_obj.write(self.turn_off_str + self.end_str)
        self.only_write(self.turn_off_str + self.end_str)
        # remember to change the state variable so change type won't have error
        self.state_o = 'off'

    def chamber_set(self, tset0):
        # turn off the load is independent command,
        # turn on the load is integrated in the change of command
        if tset0 > self.temp_H_limt:
            tset0 = self.temp_H_limt
            # also send the error message?
            # or the program can check tset from temp? to know the error
        elif tset0 < self.temp_L_limt:
            tset0 = self.temp_L_limt

        print('now is going to turn on, tset: ' + str(tset0))
        # self.inst_obj.write(self.mode_set_str + self.end_str)
        self.only_write(self.mode_set_str + self.end_str)
        # self.only_write(self.temp_set_str +
        #                 str(round(tset0, 1)) + self.end_str)
        no_save = self.query_write(self.temp_set_str +
                                   str(round(tset0, 1)) + self.end_str)
        # remember to change the state variable so change type won't have error
        self.state_o = 'on'
        self.tset_o = tset0
        # when the temperature is not ready yet, need to wait for the
        # temperature to get ready
        # to prevent read error
        time.sleep(1)
        read_temp = self.read('temp_mea')

        # variable used to break the loop in simulation mode
        # no used in the real operation
        sim_count = 0
        while abs(read_temp - self.tset_o) > self.ready_err:
            print('temp now: ' + str(read_temp) +
                  ' and target is ' + str(self.tset_o))
            print('Gary still heading over heels with Grace XD')
            # check every 10 seccond in real case
            if self.sim_inst == 1:
                time.sleep(10)
            else:
                time.sleep(0.3)
            read_temp = self.read('temp_mea')

            # to break the while loop in the simulation mode
            # setup counter for while to break
            if self.sim_inst == 0:
                sim_count = sim_count + 1
                time.sleep(0.05)
                if sim_count == 3:
                    break

        # getting out of the loop after temperature is ready
        print('temperature ok, go to the next step')
        print('Just want Grace to be happy :)')

    def ini_inst(self):
        self.chamber_off()
        # loaded the high and low limit to chamber
        self.limt_set(self.temp_H_limt_ini, self.temp_L_limt_ini)
        # loaded the acceptible error for measurement
        self.ready_err = self.ready_err_ini

        # self.chamber_set(self.tset_ini)

        pass

    def limt_set(self, h_limt, l_limt):
        self.only_write(self.temp_H_limt_set_str +
                        str(round(h_limt, 1)) + self.end_str)
        self.only_write(self.temp_L_limt_set_str +
                        str(round(l_limt, 1)) + self.end_str)
        self.temp_H_limt = h_limt
        self.temp_L_limt = l_limt

        pass

    def open_inst(self):

        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the chamber, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'chamber simulation mode object'
            pass
        # when the instrument open, need to get some intitalization

        self.ini_inst()

    def read(self, index):

        temp_string = self.query_write(
            self.temp_read_str + self.end_str)
        time.sleep(0.1)
        if self.sim_inst == 0:
            # this line is only to prevent the error from the simulation mode
            temp_string = '36.0,40.0,180.0,-50.0'
        # this delay is to prevent crash, not sure why
        self.state_o_value = []
        self.state_o_value = [float(i) for i in temp_string.split(',')]
        # update the temp measured to class memory
        self.tmea_o = self.state_o_value[0]
        # return the reated request based on index
        if index == 'temp_mea':
            return self.state_o_value[0]
        elif index == 'temp_set':
            return self.state_o_value[1]
        elif index == 'temp_H_limt':
            return self.state_o_value[2]
        elif index == 'temp_L_limt':
            return self.state_o_value[3]

    def sim_mode_out(self):
        print('tset_ini is:' + str(self.tset_ini))
        print('state_ini is:' + str(self.state_ini))
        print('temp_L_limt_ini is:' + str(self.temp_L_limt_ini))
        print('temp_H_limt_ini is:' + str(self.temp_H_limt_ini))
        print('tset_o is:' + str(self.tset_o))
        print('GP_addr_ini is:' + str(self.GP_addr_ini))
        print('state_o is:' + str(self.state_o))
        print('')
        print('state_o_value is:' + str(self.state_o_value))
        print('temp_o is:' + str(self.temp_o))
        print('temp_L_limt is:' + str(self.temp_L_limt))
        print('temp_H_limt is:' + str(self.temp_H_limt))
        print('state_o is:' + str(self.state_o))

    # consider to add the read V or read I? => but this may only for debugging

    def inst_close(self):
        print('to turn off all chamber and close GPIB')

        # 220525: ccheck if need other command

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            # GPIB device close
            time.sleep(wait_samll)
            self.inst_obj.close()

            pass
        else:
            # for the simulatiom mode of change output
            print('going to turn of the device')

            pass

    def inst_name(self):
        # get the insturment name
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.query_write(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'chamber is in sim mode'

            pass

        return self.in_name


class Rigo_DM3086 ():
    # this class used to simplify the general used function or
    # the parameter of the instrument

    # think about how to build... 220830
    # the next stage of refactoring can be
    # 1. collect the open_inst, inst_name, query_write, only_write ...
    # 2. the sharing control variable: GPIB_address, simulation mode ...

    def __init__(self, mea_v_res0, max_mea_v0, mea_i_res0, max_mea_i0, GP_addr0):
        '''
        please make sure the command is set to aligent 34401 for the regio meter
        '''

        # assign the variable into object for saving
        self.mea_v_res_ini = mea_v_res0
        self.max_mea_v_ini = max_mea_v0
        self.mea_i_res_ini = mea_i_res0
        self.max_mea_i_ini = max_mea_i0
        self.GP_addr_ini = GP_addr0
        # better to define all the parameter at the initialization
        # to prevent error at runt time or the simulation mode
        self.cmd_str_mea_v = 0
        self.cmd_str_mea_i = 0
        self.mea_v_out = 0
        self.mea_i_out = 0
        self.cmd_str_name = 0

        # set the default impedance to 10M ohm for Rigo
        self.impedance_o = '10G'

        self.sim_inst = 1
        # simulation mode for the instrument
        # default set to high, in real mode, for the simulation mode,
        # change the control varable to 0
        # this will be put in each instrument object independently
        # and you will be able to switch to simulation mode any time you want

        pass

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the Rigo, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'Rigo simulation mode object'
            pass

        pass

    def sim_mode_out(self):
        # this sub used to check the variable in therterminal when call function

        pass

    def inst_close(self):
        print('to turn off all output and close GPIB device')

        if self.sim_inst == 1:
            # turn off the output if needed

            # GPIB device close
            time.sleep(wait_samll)
            self.inst_obj.close()

            pass
        else:
            # for the simulatiom mode of change output
            print('all the channel turn off now')

            pass
        pass

    def inst_name(self):
        # get the insturment name
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            # self.in_name = self.inst_obj.query(self.cmd_str_name)
            self.in_name = self.query_write(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'instrument is in sim mode'

            pass

        return self.in_name

    def mea_i2(self, i_max='Auto', i_level='Auto'):
        '''
        give max value auto to prevent default effect \n
        0 to use the default on excel
        '''

        if i_max == 0:
            # if there are no vmax setting input, set to default when meter is defined
            i_max = self.max_mea_i_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass
        if i_level == 0:
            # if there are no vmax setting input, set to default when meter is defined
            i_level = self.mea_i_res_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass

        # define command string
        self.cmd_str_mea_i = (
            'MEASure:CURR:DC? ' + str(i_max))
        # save result to below
        # self.mea_i_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_i_out = self.mea_i_out + 1
        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            # self.mea_i_out = self.inst_obj.query(self.cmd_str_mea_i)
            self.mea_i_out = self.query_write(self.cmd_str_mea_i)
            print('I measure V2 is: ' + self.mea_i_out)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('measure current with below GPIB string (mea_I2)')
            print(str(self.cmd_str_mea_i))
            self.mea_i_out = self.mea_i_out + 1
            print(self.mea_i_out)

            pass
        # return back and save in object
        return str(self.mea_i_out)

    def mea_v2(self, v_max='Auto', v_level='Auto', impedance=0):
        '''
        give max value auto to prevent default effect \n
        0 to use the default on excel
        34401 can also operate without level, only max(range)
        '''
        # 220520: if input 0, it will be same with the original version
        # definie the command string and send the command string out to GPIB
        if v_max == 0:
            # if there are no vmax setting input, set to default when meter is defined
            v_max = self.max_mea_v_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass
        if v_level == 0:
            # if there are no vmax setting input, set to default when meter is defined
            v_level = self.mea_v_res_ini
        else:
            # if there are vmax setting input, follow the parameter input
            pass

        # since don't know how to change inpedance, disable impedance function
        # if impedance != 0:
        #     # update and record final impedance
        #     self.impedance_set(impedance)

        self.cmd_str_mea_v = (
            'MEASure:VOLT:DC? ' + str(v_max))
        # save the result to below variable
        # self.mea_v_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_v_out = self.mea_v_out + 1
        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            # self.mea_v_out = self.inst_obj.query(self.cmd_str_mea_v)
            self.mea_v_out = self.query_write(self.cmd_str_mea_v)
            print('V measure V2 is: ' + self.mea_v_out)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('measure votlage with below GPIB string (mea_V2)')
            print(str(self.cmd_str_mea_v))
            self.mea_v_out = self.mea_v_out + 1
            print(self.mea_v_out)

            pass
        # return the result back to main program, should be able to access again in this object
        return str(self.mea_v_out)

    # def impedance_set(self, impedance='10G'):
    #     '''
    #     the impedance selection will be from 10M to 10G, default
    #     is set to 10G and can be change by call this function, 10G for
    #     default is lesser loading effect
    #     '''
    #     self.impedance_o = impedance
    #     self.cmd_str_impedance = 'MEASure:VOLT:DC:IMPEdance ' + \
    #         str(self.impedance_o)

    #     res_str = self.query_write(self.cmd_str_impedance)
    #     print('the impedance setting become: ' + res_str)

    #     pass

    def query_write(self, cmd, dly=0):
        '''
        send the query command into the instrument object,\n
        cmd is the command string, dly is the delay time(default=0);
        which also include the pring command string
        '''
        # send the command and print the string
        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode

            self.return_str = self.inst_obj.query(cmd, dly)

            pass
        else:
            # for the simulatiom mode of change output
            print('query write for sub program')
            self.return_str = self.return_str + 1

            pass
        self.last_cmd = cmd
        print('command string send: ' + cmd)

        return str(self.return_str)

    def only_write(self, cmd):
        '''
        send the write command into the instrument object,\n
        cmd is the command string, dly is the delay time(default=0);
        which also include the pring command string
        '''
        # send the command and print the string

        if self.sim_inst == 1:
            # 220830 update for the independent simulation mode
            self.inst_obj.write(cmd)

            pass
        else:
            # for the simulatiom mode of change output
            print('only write for sub program')

            pass

        self.last_cmd = cmd
        print('command string send: ' + cmd)


class inst_obj_gen_sub ():
    # this class used to simplify the general used function or
    # the parameter of the instrument

    # think about how to build... 220830
    # the next stage of refactoring can be
    # 1. collect the open_inst, inst_name, query_write, only_write ...
    # 2. the sharing control variable: GPIB_address, simulation mode ...

    def __init__(self, GP_addr0):
        # assign the variable into object for saving

        self.GP_addr_ini = GP_addr0

        self.sim_inst = 1
        # simulation mode for the instrument
        # default set to high, in real mode, for the simulation mode,
        # change the control varable to 0
        # this will be put in each instrument object independently
        # and you will be able to switch to simulation mode any time you want

        pass

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the power supply, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'power supply simulation mode object'
            pass

        pass

    def sim_mode_out(self):
        # this sub used to check the variable in therterminal when call function

        pass

    def inst_close(self):
        print('to turn off all output and close GPIB device')

        if self.sim_inst == 1:
            # turn off the output if needed

            # GPIB device close
            time.sleep(wait_samll)
            self.inst_obj.close()

            pass
        else:
            # for the simulatiom mode of change output
            print('all the channel turn off now')

            pass
        pass

    def inst_name(self):
        # get the insturment name
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst_obj.query(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'instrument is in sim mode'

            pass

        return self.in_name

    # end of the example class of instrument object
    pass


pass


if __name__ == '__main__':
    # add more if selection for different instrument testing
    inst_test_ctrl = 5
    # 0 => power supply, LPS505N; 1 => meter, 34460; 2 => chroma 63600
    # 3 => source meter 2440
    # 4 => chamber su242
    # 5 => DM3068 meter

    # only run the code below when this is main program
    # can used for the testing of import, otherwise it will
    # run all the wat out after import(see strange test action)

    if inst_test_ctrl == 0:

        # power supply application
        PWR_supply1 = LPS_505N(0, 0, 1, 1, 'off')

        PWR_supply1.sim_inst = 0
        # simulation control for the power supply

        # open the GPIB device from resource manager, need to add after object is define
        PWR_supply1.open_inst()

        # check default value
        print('')
        print('check default value')
        PWR_supply1.sim_mode_out()

        # change output
        # use input to stop and key in value
        a = input()
        print('')
        print('change output')
        PWR_supply1.chg_out(3.7, 1, 1, 'on')
        # turn on channel 1 with V=3.7V and I=1A
        PWR_supply1.sim_mode_out()

        # single change V and I
        a = input()
        print('')
        print('single change for V and I')
        PWR_supply1.change_V(3.8, 3)
        PWR_supply1.sim_mode_out()

        a = input()

        PWR_supply1.change_I(2, 3)
        PWR_supply1.sim_mode_out()

        a = input()

        PWR_supply1.chg_out(4, 1.5, 1, 'off')
        PWR_supply1.sim_mode_out()

        a = input()

        PWR_supply1.iout_o = 0.5
        PWR_supply1.sim_mode_out()
        # check Iin from power supply: using iout with channel to read related IOUT
        PWR_supply1.read_iout(3)

        # close the device after experiment is finished
        PWR_supply1.inst_close()

    # meter appllication
    if inst_test_ctrl == 1:

        # definition of meter
        M1_v_in = Met_34460(0.0001, 7, 0.000001, 1, 17)
        M1_v_in.sim_inst = 0
        # simulation control for the meter

        M1_v_in.open_inst()

        print('the output of the default input')
        M1_v_in.sim_mode_out()

        a = input()

        print(' show the result of measure V 3 times, output voltage should change with measurement')
        M1_v_in.mea_v()
        M1_v_in.mea_i()

        M1_v_in.sim_mode_out()
        print('')
        print(M1_v_in.mea_v_out)
        print(M1_v_in.mea_i_out)

        a = input()

        M1_v_in.mea_v()
        M1_v_in.mea_i()
        M1_v_in.sim_mode_out()
        print('')
        print(M1_v_in.mea_v_out)
        print(M1_v_in.mea_i_out)

        a = input()

        M1_v_in.mea_v()
        M1_v_in.mea_i()
        M1_v_in.sim_mode_out()
        print('')
        print(M1_v_in.mea_v_out)
        print(M1_v_in.mea_i_out)

        # print the result got from object saving data (last measurement)

        M1_v_in.inst_close()
        # close the meter after experiment is finished

    # chroma loader 63600 application
    if inst_test_ctrl == 2:

        # here is the start point of the chroma loader testing items
        print(' the testing of chroma load sarts from here ')
        input()

        # power here set to ready by hand, only test the function of object definition and function
        # connect all the load channel to the power supply

        load1 = chroma_63600(1, 7, 'CCL')
        # object define
        # don't have status settings because all the state is set to 0

        load1.sim_inst = 0
        # simulation control for the loader

        # open the instrument
        load1.open_inst()

        # check default value
        print('default setting of object')
        load1.sim_mode_out()
        input()

        # turn on the different load
        load1.chg_out(0.2, 1, 'on')
        print('channel 1 active ')
        load1.sim_mode_out()
        input()

        load1.chg_out(0.03, 2, 'on')
        print('channel 2 active ')
        load1.sim_mode_out()
        input()

        load1.chg_out(0.1, 3, 'on')
        print('channel 3 active ')
        load1.sim_mode_out()
        input()

        load1.chg_out(0.02, 4, 'on')
        print('channel 4 active ')
        load1.sim_mode_out()
        input()

        # start to toggling the loading

        load1.chg_out(0.2, 1, 'off')
        print('turn off channel 1, only off, not change setting')
        load1.sim_mode_out()
        input()

        load1.chg_out(0.2, 1, 'on')
        print('turn on channel 1 again')
        load1.sim_mode_out()
        input()

        # single channel change loading current during on state

        load1.chg_out(0.1, 1, 'on')
        print('change the load current, CH1')
        load1.sim_mode_out()
        input()

        load1.chg_out(0.05, 4, 'on')
        print('change the load current, CH4')
        load1.sim_mode_out()
        input()

        # break point check when is the command active

        # testing for reading the voltage and current
        # to check the voltage, you can compare with the meter output

        i_check = load1.read_iout(1)

        print('the Iout of ch1 is' + i_check)
        load1.sim_mode_out()
        input()

        v_check = load1.read_vout(1)
        print('the Vout of ch1 is is' + v_check)
        load1.sim_mode_out()
        input()

        # testing for change mode

        load1.chg_mode(1, 'CCH')
        print('change the mode to CCH, CH1')
        load1.sim_mode_out()
        input()

        # re-turn on the load and change mode again
        load1.chg_out(0.5, 1, 'on')

        load1.chg_mode(1, 'CCM')
        print('change the mode to CCM, CH1')
        load1.sim_mode_out()
        input()

        # re-turn on the load and change no CC mode to test the error correction
        load1.chg_out(0.25, 1, 'on')

        load_temp = load1.chg_mode(1, 'CRM')
        print('change the mode to CRM, CH1')
        print('check if the simukation output is back to CCH')
        print('error status' + str(load_temp))
        load1.error_clear()
        load_temp = load1.errflag
        print('error status2' + str(load_temp))

        load1.sim_mode_out()
        input()

        # check the name of load
        load_temp = load1.inst_name()
        print('name of the instrument: ' + load_temp)
        input()

        # check the error response of the chg_out finction

        load1.chg_mode(1, 'CCL')
        load_temp = load1.chg_out(0.5, 1, 'on')
        print('error flag')
        print(load_temp)
        load1.error_clear()
        load_temp = load1.errflag
        print('after clear error')
        print(load_temp)
        load1.sim_mode_out()
        input()

        # testing of the close GPIB function
        load1.inst_close()

        print('testing of load is end')

    # testing for the source meter
    if inst_test_ctrl == 3:
        print('testing of the source meter starts from here')

        load_src = Keth_2440(0, 0, 24, 'off', 'CURR', 15)
        # default V and I set better set to 0
        load_src.sim_mode_out()
        print('state1')
        # after the initialization, default state ready
        input()

        load_src.sim_inst = 0
        # simulation control for the source meter

        load_src.open_inst()
        read_res = load_src.inst_name()
        print('the name is')
        print(read_res)
        print('state2')
        input()

        # loading and initialize the output parameter
        load_src.sim_mode_out()
        print('check initial parameter')
        print('state3')
        input()

        load_src.change_I(0.1, 'keep')
        # change the setting of the current and keep original state
        load_src.load_on()
        # turn the load on
        read_res = load_src.read('CURR')
        print(read_res)
        # print the read result to check
        print('should be current here? ')
        print('state4')
        input()

        load_src.change_V(16, 'keep')
        # change the clamp V and keep output on
        load_src.sim_mode_out()
        print('check the output state and votlage clamp setting here and press enter to continue')
        print('state4.5')
        input()

        load_src.change_I(0.2, 'keep')
        # change the output current during not dropping the load
        load_src.sim_mode_out()
        print('state5')
        print(
            'check the output state and curent 0.2 setting here and press enter to continue')
        read_res = load_src.read('CURR')
        print(read_res)
        read_res = load_src.read('VOLT')
        print(read_res)
        # print the read result to check
        print('should be current here? ')
        input()

        load_src.change_I(0.3, 'off')
        # change the setting after turn off first
        load_src.sim_mode_out()
        read_res = load_src.read('CURR')
        print('state6')
        print(read_res)
        read_res = load_src.read('VOLT')
        print(read_res)
        # print the read result to check
        print('should be current here? ')
        input()

        load_src.load_on()
        # turn the load on again after checking the simulation out

        # try to change type during operation, can't work with
        load_src.change_type('VOLT', 0.5)
        # change type to voltage source, the clamp setting will be current
        # no need to call the initialize since it's already include in the
        # change type sub-program
        load_src.sim_mode_out()
        # check if the type be change by sub, no going to work during operation
        print('change state during on, no used')
        print('state7')
        input()

        load_src.load_off()
        # turn the load off and change type again
        load_src.change_type('VOLT', 0.5)
        # change type to voltage source, the clamp settingwill be current
        load_src.change_V(1.5, 'off')
        load_src.sim_mode_out()
        # check if the type be change by sub, should be able to when load is already off
        print('change state during off, success')
        print('state8')
        input()

        load_src.load_on()
        # now is become voltage source an turn on
        load_src.sim_mode_out()
        read_res = load_src.read('VOLT')
        print(read_res)
        read_res = load_src.read('CURR')
        print(read_res)
        # print the read result to check
        print('should be voltage here? ')
        print('state9')
        input()

        # change I clamp during operation of votlage source
        load_src.change_I(0.6, 'keep')
        # change without turn off the load
        load_src.sim_mode_out()
        read_res = load_src.read('VOLT')
        print(read_res)
        read_res = load_src.read('CURR')
        print(read_res)
        # print the read result to check
        print('should be voltage here? ')
        print('state10')
        input()

        load_src.change_I(0.3, 'off')
        # change the current clamp again and turn off load before chagne
        load_src.sim_mode_out()
        print('change current to 0.3 with load turn off')
        print('state11')
        input()

        load_src.load_on()
        # turn the load back on after check the simulation result

        # start to check the output votlage change here
        load_src.change_V(3, 'keep')
        read_res = load_src.read('VOLT')
        print(read_res)
        read_res = load_src.read('CURR')
        print(read_res)

        print('change votlage to 3 and keep turn on')
        print('state12')
        input()

        # change the output voltage from 0V to 3V without turn off
        load_src.change_type('CURR', 1)
        load_src.sim_mode_out()

        print('should not change output, check result')
        print('state13')
        input()

        load_src.change_V(5, 'off')
        # should turn off before change command
        # command update in simulation result
        load_src.sim_mode_out()
        print('state13.5')
        input()

        load_src.load_on()
        read_res = load_src.read('VOLT')
        print(read_res)
        read_res = load_src.read('CURR')
        print(read_res)
        print('state14')
        input()

        load_src.inst_close()
        # turn the load off and close GPIB object after operation finished

    if inst_test_ctrl == 4:
        # the chamber related testing
        temp_res = [0, 0, 0, 0]
        temp_res2 = ''
        temp_string = '36.0,40.0,180.0,-50.0'
        float_list = []
        float_list = [float(i) for i in temp_string.split(',')]
        print(float_list)

        cham = chamber_su242(25, 15, 'off', -45, 185, 0)
        cham.sim_mode_out()
        cham.sim_inst = 0
        # simulation control for the chamber

        input()
        cham.open_inst()
        cham.ini_inst()
        cham.limt_set(100, 0)
        cham.sim_mode_out()
        input()

        cham.chamber_set(-10)
        cham.sim_mode_out()
        input()

        temp_res = cham.read('temp_mea')
        print(str(temp_res))
        temp_res = cham.read('temp_set')
        print(str(temp_res))
        temp_res = cham.read('temp_H_limt')
        print(str(temp_res))
        temp_res = cham.read('temp_L_limt')
        print(str(temp_res))

        input()

        cham.ini_inst()
        cham.sim_mode_out()
        input()

        pass

    # DM3068 testing
    if inst_test_ctrl == 5:

        rigo_m = Rigo_DM3086(0.0001, 7, 0.000001, 1, 20)
        rigo_m.open_inst()
        temp_str = rigo_m.inst_name()
        print(temp_str)

        temp_str = rigo_m.mea_v2()
        print(temp_str)
        temp_str = rigo_m.mea_i2()
        print(temp_str)
        # temp_str = rigo_m.impedance_o
        temp_str = rigo_m.mea_v2(0)
        print(temp_str)
        # temp_str = rigo_m.mea_v2(30, 0.000001, '10M')
        temp_str = rigo_m.mea_i2(0)
        print(temp_str)

        # rigo_m.impedance_set('10M')
        # temp_str = rigo_m.impedance_o
        # print(temp_str)

        pass
