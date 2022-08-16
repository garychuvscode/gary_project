import pyvisa

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

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        self.power = rm.open_resource(
            'GPIB0::' + str(self.GP_addr_ini) + '::INSTR')

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
            str(self.act_ch_o) + ":" + str(self.state_o)
        # using normal(independent mode, not serial)
        self.cmd_str_out_mode = "OUT: normal"
        # setting string for the V and I
        self.cmd_str_V = ("PROG:VSET" + str(self.act_ch_o) +
                          ":" + str(self.vset_o))
        self.cmd_str_I = ("PROG:ISET" + str(self.act_ch_o) +
                          ":" + str(self.iset_o))
        # all command string using self because you can reference after the function is over,
        # it will left in the object, not disappear

        # write the command string for change power supply output

        # self.power.write(self.cmd_str_out_sw)
        # self.power.write(self.cmd_str_out_mode)
        # self.power.write(self.cmd_str_V)
        # self.power.write(self.cmd_str_I)

    # this function used to fast change voltage setting

    def change_V(self, vset2):
        self.vset_o = vset2
        self.cmd_str_V = ("PROG:VSET" + str(self.act_ch_o) +
                          ":" + str(self.vset_o))
        # must change both voltage and current together for every update
        # so the power supply will change output
        # self.power.write(self.cmd_str_V)
        # self.power.write(self.cmd_str_I)

    # this function used to fast change current setting

    def change_I(self, iset2):
        self.iset_o = iset2
        self.cmd_str_I = ("PROG:ISET" + str(self.act_ch_o) +
                          ":" + str(self.iset_o))
        # must change both voltage and current together for every update
        # so the power supply will change output
        # self.power.write(self.cmd_str_V)
        # self.power.write(self.cmd_str_I)

    # this function only used to check the simulation mode ouptput

    def read_iout(self):
        # self.iout_o = 0
        # not to initialize variable outside of the constructor
        # this will cause the erase of previous result
        self.cmd_str_iout = 'IOUT' + str(self.act_ch_o) + '?'

        # self.iout_o = self.power.query(self.cmd_str_iout)

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

        # self.power.close()


PWR_supply1 = LPS_505N(0, 0, 1, 1, 'off')
# open the GPIB device from resource manager, need to add after object is define
# PWR_supply1.open_inst()

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

PWR_supply1.change_I(2)
PWR_supply1.sim_mode_out()

PWR_supply1.chg_out(4, 1.5, 1, 'off')
PWR_supply1.sim_mode_out()

PWR_supply1.iout_o = 0.5
PWR_supply1.sim_mode_out()
# check Iin from power supply: using iout with channel to read related IOUT
PWR_supply1.read_iout()


# close the device after experiment is finished
PWR_supply1.inst_close()


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

    def open_instr(self):
        self.meter = rm.open_resource(
            'GPIB0::' + str(self.GP_addr_ini) + '::INSTR')

    def mea_v(self):
        # definie the command string and send the command string out to GPIB
        self.cmd_str_mea_v = (
            'MEASure:VOLT:DC? ' + str(self.max_mea_v_ini) + ',' + str(self.mea_v_res_ini))
        # save the result to below variable
        # self.mea_v_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        self.mea_v_out = self.mea_v_out + 1
        # self.mea_v_out = self.meter.query(self.cmd_str_mea_v)

        # return the result back to main program, should be able to access again in this object
        return self.mea_v_out

    def mea_i(self):
        # define command string
        self.cmd_str_mea_i = (
            'MEASure:CURR:DC? ' + str(self.max_mea_i_ini) + ',' + str(self.mea_i_res_ini))
        # save result to below
        # self.mea_i_out = 0
        # not to set to 0, definition create at the initialization function, so can keep the old result from last measurement
        self.mea_i_out = self.mea_i_out + 1
        # self.mea_i_out = self.meter.query(self.cmd_str_mea_i)

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
        # print(self.meter)
        print(self.cmd_str_mea_v)
        print(self.cmd_str_mea_i)
        print('')
        print(self.mea_v_out)
        print(self.mea_i_out)

    def inst_close(self):
        print('to turn off all output and close GPIB device')
        # change all the output to 0V and 0A, for channel 1 to 3
        # self.meter.close()
# meter appllication


# definition of meter
M1_v_in = Met_34460(0.0001, 7, 0.000001, 1, 20)
# M1_v_in.open_instr()

print('the output of the default input')
M1_v_in.sim_mode_out()
print(' show the result of measure V 3 times, output voltage should change with measurement')
M1_v_in.mea_v()
M1_v_in.mea_i()

M1_v_in.sim_mode_out()
print('')
print(M1_v_in.mea_v_out)

M1_v_in.mea_v()
M1_v_in.mea_i()
M1_v_in.sim_mode_out()
print('')
print(M1_v_in.mea_v_out)

M1_v_in.mea_v()
M1_v_in.mea_i()
M1_v_in.sim_mode_out()
print('')
print(M1_v_in.mea_v_out)
print(M1_v_in.mea_i_out)

# print the result got from object saving data (last measurement)

M1_v_in.inst_close()
# close the meter after experiment is finished
