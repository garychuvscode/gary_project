from os import supports_bytes_environ
import pyvisa
import time
wait_samll = 0.05
# 100ms of waiting time for the separation of GPIB command
# maybe the instrument need the delay time
rm = pyvisa.ResourceManager()


class LPS_505N:
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

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        self.power = rm.open_resource(
            'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        time.sleep(wait_samll)

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

        self.power.write(self.cmd_str_out_mode)
        self.power.write(self.cmd_str_V)
        self.power.write(self.cmd_str_I)
        self.power.write(self.cmd_str_out_sw)
        time.sleep(wait_samll)

    # this function used to fast change voltage setting

    def change_V(self, vset2):
        self.vset_o = vset2
        self.cmd_str_V = ("PROG:VSET" + str(int(self.act_ch_o)) +
                          ":" + str(self.vset_o))
        # must change both voltage and current together for every update
        # so the power supply will change output
        self.power.write(self.cmd_str_V)
        self.power.write(self.cmd_str_I)
        # when the source is already on, need to have the turn on command to update the final
        # setting to output, to prevent wrong behavior of the power supply
        # (only change the output voltage but not the current)
        # use another command to update at the same time
        self.power.write(self.cmd_str_out_sw)
        time.sleep(wait_samll)

    # this function used to fast change current setting

    def change_I(self, iset2):
        self.iset_o = iset2
        self.cmd_str_I = ("PROG:ISET" + str(int(self.act_ch_o)) +
                          ":" + str(self.iset_o))
        # must change both voltage and current together for every update
        # so the power supply will change output
        self.power.write(self.cmd_str_V)
        self.power.write(self.cmd_str_I)
        # when the source is already on, need to have the turn on command to update the final
        # setting to output, to prevent wrong behavior of the power supply
        # (only change the output voltage but not the current)
        # use another command to update at the same time
        self.power.write(self.cmd_str_out_sw)
        time.sleep(wait_samll)

    # this function only used to check the simulation mode ouptput

    def read_iout(self):
        # self.iout_o = 0
        # not to initialize variable outside of the constructor
        # this will cause the erase of previous result
        self.cmd_str_iout = 'IOUT' + str(int(self.act_ch_o)) + '?'

        self.iout_o = self.power.query(self.cmd_str_iout)
        time.sleep(wait_samll)
        return self.iout_o

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
        # change all the output to 0V and 0A, for channel 1 to 3
        self.chg_out(0, 0, 1, 'off')
        self.chg_out(0, 0, 2, 'off')
        self.chg_out(0, 0, 3, 'off')
        # GPIB device close
        time.sleep(wait_samll)
        self.power.close()

    def inst_name(self):
        self.cmd_str_name = "*IDN?"
        self.in_name = self.power.query(self.cmd_str_name)
        time.sleep(wait_samll)
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

        self.power.write(self.cmd_str_ovp)
        self.power.write(self.cmd_str_ocp)
        time.sleep(wait_samll)


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

    def open_instr(self):
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        self.meter = rm.open_resource(
            'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        time.sleep(wait_samll)

    def mea_v(self):
        # definie the command string and send the command string out to GPIB
        self.cmd_str_mea_v = (
            'MEASure:VOLT:DC? ' + str(self.max_mea_v_ini) + ',' + str(self.mea_v_res_ini))
        # save the result to below variable
        # self.mea_v_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_v_out = self.mea_v_out + 1
        self.mea_v_out = self.meter.query(self.cmd_str_mea_v)
        time.sleep(wait_samll)
        # return the result back to main program, should be able to access again in this object
        return self.mea_v_out

    def mea_i(self):
        # define command string
        self.cmd_str_mea_i = (
            'MEASure:CURR:DC? ' + str(self.max_mea_i_ini) + ',' + str(self.mea_i_res_ini))
        # save result to below
        # self.mea_i_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        # self.mea_i_out = self.mea_i_out + 1
        self.mea_i_out = self.meter.query(self.cmd_str_mea_i)
        time.sleep(wait_samll)
        # return back and save in object
        return self.mea_i_out

    def sim_mode_out(self):
        # simulation output mode to check result in terminal
        print(self.mea_v_res_ini)
        print(self.max_mea_v_ini)
        print(self.mea_i_res_ini)
        print(self.max_mea_i_ini)
        print(self.GP_addr_ini)
        print('')
        print(self.meter)
        print(self.cmd_str_mea_v)
        print(self.cmd_str_mea_i)
        print('')
        print(self.mea_v_out)
        print(self.mea_i_out)

    def inst_close(self):
        print('to turn off all output and close GPIB device')
        time.sleep(wait_samll)
        # change all the output to 0V and 0A, for channel 1 to 3
        self.meter.close()

    def inst_name(self):
        self.cmd_str_name = "*IDN?"
        self.in_name = self.meter.query(self.cmd_str_name)
        time.sleep(wait_samll)
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
        self.i_sel_ch = [0, 0, 0, 0]
        # array of current setting, map number = number -1
        # self.i_sel_ch1 = 0
        # self.i_sel_ch2 = 0
        # self.i_sel_ch3 = 0
        # self.i_sel_ch4 = 0
        self.mode_o = self.mode_ini
        # only change mode o, not change mode_ini, data loaded from mode o

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
        self.cmd_str_name = 0
        self.in_name = 0

        # other parameter needed in loader class
        self.v_out = 0
        self.i_out = 0

        # loader object definition
        self.loader = 0
        # other definition
        self.errflag = 0
        # error flag indicate => 0 is ok and 1 is error

    # need to watch out if the power limit, current limit need to be set and config through
    # the different mode of the load setting, to prevent crash of the load during auto testing
    # the accuracy and current operating range will also be different for the different mode

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        # here is to define object map to the loader
        # conntect to the related GPIB address
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        self.loader = rm.open_resource(
            'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        time.sleep(wait_samll)

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
        self.cmd_str_I_load = "curr:stat:L1 " + \
            str(self.i_sel_ch[int(self.act_ch_o) - 1])
        print(self.cmd_str_I_load)
        # update the status of on and off
        self.cmd_str_status = ("Load " + self.state_o[int(self.act_ch_o) - 1])
        print(self.cmd_str_status)
        # all command string using self because you can reference after the function is over,
        # it will left in the object, not disappear

        # write the command string for change power supply output
        # writring sequence: channel => current setting => status update
        # check if it changes like power supply? need to update status to refresh command
        self.loader.write(self.cmd_str_ch_set)
        self.loader.write(self.cmd_str_mode_set)
        self.loader.write(self.cmd_str_I_load)
        self.loader.write(self.cmd_str_status)
        # add the break point here to double if the command update based on the write command of status
        time.sleep(wait_samll)

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

        # string write and action
        self.loader.write(self.cmd_str_ch_set)
        self.v_out = self.loader.query(self.cmd_str_V_read)
        time.sleep(wait_samll)
        return self.v_out

    # this function used to feedback the output current measurement

    def read_iout(self, act_ch1):
        # also need to input the channel to check, and it will update the final channel
        self.act_ch_o = act_ch1

        # string setting and update
        self.cmd_str_ch_set = "CHAN " + str(int(self.act_ch_o))
        self.cmd_str_V_read = ("MEAS:CURR?")

        # string write and action
        self.loader.write(self.cmd_str_ch_set)
        self.i_out = self.loader.query(self.cmd_str_V_read)
        time.sleep(wait_samll)
        return self.i_out

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
        print(self.loader)
        print(self.errflag)

    # consider to add the read V or read I? => but this may only for debugging

    # to close all the channel and turn off the GPIB connection, need to open again

    def inst_close(self):
        print('to turn off all output and close GPIB device')
        # change all the output to 0V and 0A, for channel 1 to 3
        self.chg_out(0, 1, 'off')
        self.chg_out(0, 2, 'off')
        self.chg_out(0, 3, 'off')
        self.chg_out(0, 4, 'off')
        # GPIB device close
        time.sleep(wait_samll)
        self.loader.close()

    # used to get the instrument name and return to main program
    def inst_name(self):
        self.cmd_str_name = "*IDN?"
        self.in_name = self.loader.query(self.cmd_str_name)
        time.sleep(wait_samll)
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


if __name__ == '__main__':
    # add more if selection for different instrument testing
    inst_test_ctrl = 2
    # 0 => power supply, LPS505N; 1 => meter, 34460; 2 => chroma 63600

    # only run the code below when this is main program
    # can used for the testing of import, otherwise it will
    # run all the wat out after import(see strange test action)

    if inst_test_ctrl == 0:

        # power supply application
        PWR_supply1 = LPS_505N(0, 0, 1, 1, 'off')
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
        PWR_supply1.change_V(3.8)
        PWR_supply1.sim_mode_out()

        a = input()

        PWR_supply1.change_I(2)
        PWR_supply1.sim_mode_out()

        a = input()

        PWR_supply1.chg_out(4, 1.5, 1, 'off')
        PWR_supply1.sim_mode_out()

        a = input()

        PWR_supply1.iout_o = 0.5
        PWR_supply1.sim_mode_out()
        # check Iin from power supply: using iout with channel to read related IOUT
        PWR_supply1.read_iout()

        # close the device after experiment is finished
        PWR_supply1.inst_close()

    # meter appllication
    if inst_test_ctrl == 1:

        # definition of meter
        M1_v_in = Met_34460(0.0001, 7, 0.000001, 1, 22)
        M1_v_in.open_instr()

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
