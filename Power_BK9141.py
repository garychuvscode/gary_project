import pyvisa
import re
import time
import locale as lo

from GInst import *
import time

wait_samll = 0.05
wait_time = 0.2

rm = pyvisa.ResourceManager()


class Power_BK9141(GInst):
    '''
    Class library from Geroge is channel based instrument, need to define by channel
    not a single instrument \n
    for parallel operation, need to setup by hand and use the channel 1 as control window
    '''

    def __init__(self, link='', ch='CH1', excel0=0, ini=0, GP_addr0=0, main_off_line0=0):
        super().__init__()
        '''
        ini = 0 is not to use geroge's open instrument, and it's been define as single channel
        add other function of operate as a whole instrument

        '''

        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import parameter_load_obj as par
            # for the jump out window

            # initial the object and set to simulation mode

            # using the main control book as default
            excel0 = par.excel_parameter('obj_main')
            # ======== only for object programming

        self.excel_s = excel0
        self.cmd_str_name = ''
        # the information for GPIB resource manager
        if GP_addr0 == 0:
            # which means no GPIB address input for object
            self.GP_addr_ini = self.excel_s.pwr_bk_addr
        else:
            self.GP_addr_ini = GP_addr0

        if self.GP_addr_ini != 100 and main_off_line0 == 0:
            self.sim_inst = 1
            link = 'GPIB0::' + str(int(self.excel_s.pwr_bk_addr)) + '::INSTR'
        else:
            self.sim_inst = 0

        # self.open_inst()

        if ini == 1:
            # move the rm of Geroge to the top, since there will be other way to open instrument
            # rm = pyvisa.ResourceManager()

            self.link = link
            # link is the GPIB string for pyvisa resource manager
            self.ch = ch
            # ch is the channel index for in below dictionary
            self.chConvert = {'CH1': '0', 'CH2': '1',
                              'CH3': '2', 'ALL': ':ALL'}

            try:
                self.inst = rm.open_resource(link)
                self.inst.read_termination = '\n'
                self.inst.write_termination = '\n'
                self.inst.baud_rate = 38400
                self.inst.timeout = 500

            except Exception as e:
                raise Exception(
                    f'<>< Power_BK9141 ><> open 3-CH Power Fail {str(e)}!')

            idn = self.inst.query('*IDN?')

            if '9141' not in idn:
                raise Exception(
                    f'<>< Power_BK9141 ><> "9141" not fit in "{idn}"!')

        # 221221 add the voltage and current limit for the BK9141
        # since it's for HV buck, default set to 28V first
        self.max_v = 28
        self.max_i = 3

    # @GInstSetMethod(unit = 'V')

    def setVoltage(self, voltage, ch_str=0):
        """
        power.setVoltage(voltage) -> None
        ================================================================
        [power(channel) set Voltage]
        :param voltage:
        :return: None. \n
        ch_str can set other channel: CHx
        """
        if voltage is None:
            return

        if ch_str == 0:
            self.inst.write(f'INST {self.chConvert[self.ch]}')
        else:
            self.inst.write(f'INST {str(ch_str)}')
        self.inst.write(f'VOLT {abs(voltage):3g}')

        pass

    # @GInstSetMethod(unit = 'A')

    def setCurrent(self, current, ch_str=0):
        """
        power.setCurrent(current) -> None
        ================================================================
        [power(channel) set Current]
        :param current:
        :return: None.
        """
        if current is None:
            return

        if ch_str == 0:
            self.inst.write(f'INST {self.chConvert[self.ch]}')
        else:
            self.inst.write(f'INST {str(ch_str)}')
        self.inst.write(f'CURR {current:3g}')

        pass

    def change_I(self, iset1, ch=0):
        # mapped to the change_I function of precious power supply
        iset1 = float(iset1)

        # 221221 update => refer to chg_out function
        if ch != 0 and iset1 <= self.max_i:
            self.setCurrent(iset1, 'CH' + str(ch))
        else:
            print('no channel or higher than max_i')

    def ov_oc_set(self, max_v0, max_i0):
        '''
        update the default vin_max(28V) and iin_max(3A), mainly for high V buck
        '''
        self.max_v = max_v0
        self.max_i = max_i0

        pass

    # @GInstOnMethod()

    def outputON(self, ch_str=0):
        """
        power.outputON() -> None
        ================================================================
        [power(channel) output ON]
        :param None:
        :return: None.
        """
        if ch_str == 0:
            self.inst.write(f'INST {self.chConvert[self.ch]}')
        else:
            self.inst.write(f'INST {str(ch_str)}')
        self.inst.write(f'OUTP 1')

    # @GInstOffMethod()

    def outputOFF(self, ch_str=0):
        """
        power.outputOFF() -> None
        ================================================================
        [power(channel) output OFF]
        :param None:
        :return: None.
        """
        if ch_str == 0:
            self.inst.write(f'INST {self.chConvert[self.ch]}')
        else:
            self.inst.write(f'INST {str(ch_str)}')
        self.inst.write(f'OUTP 0')

    # @GInstGetMethod(unit = 'V')

    def measureVoltage(self, ch_str=0):
        """
        power.measureVoltage() -> Voltage
        ================================================================
        [power(channel) measure Voltage]
        :param None:
        :return: Voltage. \n
        ch_str = 'CH1'
        """
        try:
            if ch_str == 0:
                self.inst.write(f'INST {self.chConvert[self.ch]}')
            else:
                self.inst.write(f'INST {str(ch_str)}')
            valstr = self.inst.query(f'MEAS:SCAL:VOLTage:DC?')
        except pyvisa.errors.VisaIOError:
            # here is the old ersion from geroge
            valstr = self.inst.query(f'MEAS:SCAL:VOLTage:DC?')
            self.inst.query(f'*CLS?')

            # # 221221: change the exception operation to clear fault first
            # # and read again; the other one is change to write since
            # # there may not be return for '*CLS' (clear status and errors)
            # self.inst.query(f'*CLS')
            # valstr = self.inst.query(f'MEAS:SCAL:VOLTage:DC?')

        return float(re.search(r"[-+]?\d*\.\d+|\d+", valstr).group(0))

    # @GInstGetMethod(unit = 'A')

    def measureCurrent(self, ch_str=0):
        """
        power.measureCurrent() -> Current
        ================================================================
        [power(channel) measure Current]
        :param None:
        :return: Current.\n
        ch_str = 'CH1'
        """
        try:
            if ch_str == 0:
                self.inst.write(f'INST {self.chConvert[self.ch]}')
            else:
                self.inst.write(f'INST {str(ch_str)}')
            valstr = self.inst.query(f'MEAS:SCAL:CURR:DC?')
        except pyvisa.errors.VisaIOError:
            valstr = self.inst.query(f'MEAS:SCAL:CURR:DC?')
            self.inst.query(f'*CLS?')

        return float(re.search(r"[-+]?\d*\.\d+|\d+", valstr).group(0))

    def read_iout(self, ch0=0):
        '''
        mapped with the read iout function of pwr
        '''
        ch0 = str(int(ch0))
        ch0_1 = 'CH' + ch0

        # try to wait for a moment of read current
        time.sleep(0.1)

        curr = self.measureCurrent(ch_str=ch0_1)

        return str(curr)
        pass

    def chg_out(self, vset1='NA', iset1='NA', act_ch1=1, state1='NA'):

        act_ch1 = int(act_ch1)

        if vset1 != 'NA':
            if vset1 <= self.max_v:
                self.setVoltage(vset1, 'CH' + str(act_ch1))
        if iset1 != 'NA':
            if iset1 <= self.max_i:
                self.setCurrent(iset1, 'CH' + str(act_ch1))

        if state1 != 'NA':
            if state1 == 'on':
                self.outputON('CH' + str(act_ch1))
            else:
                self.outputOFF('CH' + str(act_ch1))

        pass

    def vin_calibrate_singal_met(self, vin_ch, vin_target, met_v0, mcu0, excel0):
        '''
        vin_ch is the channel of relay
        '''
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
            vin_new = vin_target + excel0.pre_inc_vin
            while vin_diff > excel0.vin_diff_set or vin_diff < (-1 * excel0.vin_diff_set):
                vin_new = vin_new + 0.5 * (vin_target - v_res_temp_f)
                # clamp for the Vin maximum
                if vin_new > excel0.pre_vin_max:
                    vin_new = excel0.pre_vin_max

                if vin_new < 0:
                    vin_new = 0

                if vin_ch == 0:
                    # self.change_V(vin_new, excel0.relay0_ch)
                    self.chg_out(act_ch1=excel0.relay0_ch,
                                 vset1=vin_new, state1='on')
                    # send the new Vin command for the auto testing channel
                    pass

                elif vin_ch == 6:
                    # self.change_V(vin_new, excel0.relay6_ch)
                    self.chg_out(act_ch1=excel0.relay6_ch,
                                 vset1=vin_new, state1='on')
                    # change the vsetting of channel 1 (mapped in program)
                    pass

                elif vin_ch == 7:
                    # self.change_V(vin_new, excel0.relay7_ch)
                    self.chg_out(act_ch1=excel0.relay7_ch,
                                 vset1=vin_new, state1='on')
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
            v_res_temp = int(v_res_temp)
            v_res_temp = v_res_temp + 1
        # need to return the channel after the calibration is finished

        # the last measured value can also find in the meter result
        return str(v_res_temp)

    def open_inst(self):
        # maybe no need to define rm for global variable
        # global rm
        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(wait_samll)
            pass
        else:
            print('now is open the power supply, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst need to be define for the simuation mode
            self.inst = 'power supply simulation mode object'
            pass

    def inst_name(self):
        # get the insturment name
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst.query(self.cmd_str_name)
            time.sleep(wait_samll)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'pwr is in sim mode'

            pass

        return self.in_name


if __name__ == '__main__':
    # testing for the 9141
    test_mode = 1
    sim_test_set = 1

    import mcu_obj as mcu
    import inst_pkg_d as inst

    import parameter_load_obj as par
    excel_t = par.excel_parameter('970_full')

    met_v_t = inst.Met_34460(0.0001, 30, 0.000001, 2.5, 20)
    met_v_t.sim_inst = sim_test_set
    met_v_t.open_inst()

    mcu_t = mcu.MCU_control(1, 13)
    mcu_t.com_open()

    # bk_9141 = Power_BK9141(sim_inst0=sim_test_set, excel0=excel_t, addr=2)
    bk_9141 = Power_BK9141(excel0=excel_t)

    if test_mode == 0:
        # old version of the chg_out, with wrong sequence
        '''
        should follow LPS505
        def chg_out(self, vset1, iset1, act_ch1, state1):

        define in this test_mode
        def chg_out(self, act_ch1=1, vset1='NA', iset1='NA', state1='NA'):
        '''

        bk_9141.open_inst()
        bk_9141.chg_out(1, 3.7, 1, 'on')
        bk_9141.chg_out(1, 3.7, 1, 'off')
        bk_9141.chg_out(1, 3.7, 2, 'on')
        bk_9141.chg_out(1, 4, 1, 'on')

        bk_9141.vin_calibrate_singal_met(
            0, 4.5, met_v0=met_v_t, mcu0=mcu_t, excel0=excel_t)

        bk_9141.chg_out(2, 2, 1, 'on')
        bk_9141.chg_out(3, 1.5, 2, 'on')
        bk_9141.chg_out(3, 1.7, 2)
        bk_9141.chg_out(2, 1.9, 2)

        pass

    if test_mode == 1:

        bk_9141.open_inst()
        bk_9141.chg_out(3.7, 1, 1, 'on')
        bk_9141.chg_out(3.7, 1, 1, 'off')
        bk_9141.chg_out(3.7, 2, 1, 'on')
        bk_9141.chg_out(4, 1, 1, 'on')

        bk_9141.vin_calibrate_singal_met(
            0, 4.5, met_v0=met_v_t, mcu0=mcu_t, excel0=excel_t)

        bk_9141.chg_out(2, 1, 2, 'on')
        bk_9141.chg_out(1.5, 2, 3, 'on')
        bk_9141.chg_out(1.7, 2, 3)
        bk_9141.chg_out(1.9, 2, 2)

        a = bk_9141.read_iout(1)
        a = bk_9141.read_iout(2)
        a = bk_9141.read_iout(3)

        pass
