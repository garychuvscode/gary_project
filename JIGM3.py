
import logging
import json
import time

# from PySide6.QtCore import *  # type: ignore
# from PySide6.QtGui import *  # type: ignore
# from PySide6.QtWidgets import *  # type: ignore

from Except import *

from PyUSBManagerv4 import opendev, closedev, luacmd, listdevs


class JIGM3():

    def __init__(self, sim_mcu0=0, com_addr0=0, devpath=0, CmdSentSignal=0):

        self.DevPath = devpath
        self.handle = opendev(devpath)

        self.CmdSentSignal = CmdSentSignal

        self.Version = self.getversion()

        logging.info(f"OpenJIGM3[{self.Version}]")

        # === Gary add on for JIGM3
        '''
        The add on for initial object is for default setting of JIGM3 or the generation
        of variable
        '''

        self.sim_mcu = sim_mcu0
        self.com_addr = com_addr0

        # port number assignment
        self.i_o_port_num = {'PG': '0', 'IO': '1', 'IOx': '2', 'IOy': '3'}

        # IO number assignemnt
        self.i_o_pin_num_set = {'1': 0x1, '2': 0x2, '3': 0x4, '4': 0x8, '5': 0x10, '6': 0x20, '7': 0x40, '8': 0x80,
                                '9': 0x100, '10': 0x200, '11': 0x400, '12': 0x800, '13': 0x1000, '14': 0x2000, '15': 0x4000, '16': 0x8000, }
        self.i_o_pin_num_clr = {'1': 0xFFFE, '2': 0xFFFD, '3': 0xFFFB, '4': 0xFFF7, '5': 0xFFEF, '6': 0xFFDF, '7': 0xFFBF,
                                '8': 0xFF7F, '9': 0xFEFF, '10': 0xFDFF, '11': 0xFBFF, '12': 0xF7FF, '13': 0xEFFF, '14': 0xDFFF, '15': 0xBFFF, '16': 0x7FFF, }
        # binary number: 0b10101010

        # default state and state record of IO pin
        self.i_o_state = {'PG': 0x00, 'IO': 0x00, 'IOx': 0x00, 'IOy': 0x00}

        '''
        MSP430 related variable mapping
        '''
        # default PMIC mode is set to 1 (shut down mode)
        self.mode_set = 1

    @staticmethod
    def listdevices():
        try:
            pathm4s = listdevs((0x04d8), (0x003E), (0x4d1e55b2, 0xf16f,
                               0x11cf, 0x88, 0xcb, 0x00, 0x11, 0x11, 0x00, 0x00, 0x30))
            return pathm4s

        except Exception as e:
            logging.warning(e)
            return []

    # JIG-M4 style

    def getversion(self):
        return self.ezCommand('return mcu._version:get()')[0]

    def refresh(self):
        # self.reopen()
        pass

    def reopen(self):
        if hasattr(self, "handle") and self.handle != None:
            closedev(self.handle)

        self.handle = opendev(self.DevPath)

    def close(self):
        if hasattr(self, "handle") and self.handle != None:
            closedev(self.handle)

    # -----------------------------------------------------------------------------------------------------------------
    def retry_and_refresh(func):
        def wrapper(*args, **kwargs):
            # if (time.time() - args[0].LastTimeRefresh) > 1.0 :
            #     args[0].reopen()

            # args[0].LastTimeRefresh = time.time()
            retryed = 0
            while True:
                try:
                    return func(*args, **kwargs)
                except I2CError as e:
                    logging.warning(f"I2CError: {str(e)}")
                    raise e

                except Exception as e:
                    logging.warning(f"Retry for {str(e)}")
                    if retryed < 2:
                        # QThread.msleep(3000*retryed + 1)
                        args[0].reopen()
                        retryed += 1
                    else:
                        raise JIGError(str(e))

        return wrapper

    def assertResult(self, result):
        if not isinstance(result, list):
            raise Exception("[Result Error] Result is not list")

        if not isinstance(result[0], int):
            raise Exception(
                "[Result Error] Result[0] is not integer for Error-number")

        return result

    # -----------------------------------------------------------------------------------------------------------------
    # --  EXPORT ------------------------------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------
    @retry_and_refresh
    def ezCommand(self, command):
        """
        send command to v4 JIG
        :param command: command.
        :return: JIG result
        """
        '''
        Gary Q:
        1. what is the return value? => if sending I/O pattern from pattern gen or I2C command?
        2. what is the difference between i2c i2c_op i2c16?
        A2: i2C16 is 16 bit data length => also address?; i2C should be the normal read an write with 8bit data and
        7 bit address
        Gary note:
        1. ezcommand can be found from the V4 tool "MCU", the string needed for related operation for JIGM3
        2. pattern gen, I/O toggle and I2C command supported
        '''
        # self.CmdSentSignal.emit(command)
        logging.info(f"{command=}")

        result = json.loads(luacmd(self.handle, command))
        logging.info(f"    {result=}")

        return result

    @retry_and_refresh
    def i2c_read(self, device, regaddr, len):
        """
        read i2c
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param len: length for read.
        :return: error id, datas or error string.
        """

        '''
        return datas doen's contain register address
        '''

        cmd = f" return mcu.i2c.read(0x{device:02X}, 0x{regaddr:02X}, {len})"

        # self.CmdSentSignal.emit(cmd)

        err, datas, * \
            _ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

        match err:
            case 0:
                return datas

            case 0x03:
                raise I2CNACKError(f"{device=:02X}")
            case 0x11:
                raise I2CBusError(f"{device=:02X}")
            case _:
                raise I2CError(f"{device=:02X}")

    @retry_and_refresh
    def i2c_write(self, device, regaddr, datas):
        """
        write i2c
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param datas: list for bytes to write.
        :return: error id, datas or error string.
        """

        '''
        device address use 8 bit address, regaddr (0x9C), datas (0xFF or 0x00)-list no need index
        '''
        hexstrs = ["0x{:02X}".format(data) for data in datas]
        cmddstr = ", ".join(hexstrs)

        cmd = f" return mcu.i2c.write(0x{device:02X}, 0x{regaddr:02X}, {{{cmddstr}}})"

        # self.CmdSentSignal.emit(cmd)

        err, *_ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

        match err:
            case 0:
                return

            case 0x03:
                raise I2CNACKError(f"{device=:02X}")
            case 0x11:
                raise I2CBusError(f"{device=:02X}")
            case _:
                raise I2CError(f"{device=:02X}")

    @retry_and_refresh
    def i2c_opread(self, device, len, contAck, lastAck):
        """
        option read i2c
        :param device: I2C slave address.
        :param len: length for read.
        :param contAck: 1 for middle Ack. otherwise 0
        :param lastAck: 1 for last Ack.   otherwise 0
        :return: error id, datas or error string.
        """

        '''
        only device address, not register adress, difference between op and normal
        '''

        cmd = f' return mcu.i2c.opread(0x{device:02X}, {len}, {1 if contAck else 0}, {1 if lastAck else 0})'

        # self.CmdSentSignal.emit(cmd)

        err, datas, * \
            _ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

        match err:
            case 0:
                return datas

            case 0x03:
                raise I2CNACKError(f"{device=:02X}")
            case 0x11:
                raise I2CBusError(f"{device=:02X}")
            case _:
                raise I2CError(f"{device=:02X}")

    @retry_and_refresh
    def i2c_opwrite(self, device, datas):
        """
        option write i2c
        :param device: I2C slave address.
        :param datas: list for bytes to write.
        :return: error id, datas or error string.
        """
        hexstrs = ["0x{:02X}".format(data) for data in datas]
        cmddstr = ", ".join(hexstrs)

        cmd = f" return mcu.i2c.opwrite(0x{device:02X}, {{{cmddstr}}})"

        # self.CmdSentSignal.emit(cmd)

        err, *_ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

        match err:
            case 0:
                return

            case 0x03:
                raise I2CNACKError(f"{device=:02X}")
            case 0x11:
                raise I2CBusError(f"{device=:02X}")
            case _:
                raise I2CError(f"{device=:02X}")

    @retry_and_refresh
    def i2c_read16(self, device, regaddr, len):
        """
        read i2c
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param len: length for read.
        :return: error id, datas or error string.
        """

        cmd = f" return mcu.i2c.read16(0x{device:02X}, 0x{regaddr:04X}, {len})"

        # self.CmdSentSignal.emit(cmd)

        err, datas, * \
            _ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

        match err:
            case 0:
                return datas

            case 0x03:
                raise I2CNACKError(f"{device=:02X}")
            case 0x11:
                raise I2CBusError(f"{device=:02X}")
            case _:
                raise I2CError(f"{device=:02X}")

    @retry_and_refresh
    def i2c_write16(self, device, regaddr, datas):
        """
        write i2c
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param datas: list for bytes to write.
        :return: error id, datas or error string.
        """
        hexstrs = ["0x{:02X}".format(data) for data in datas]
        cmddstr = ", ".join(hexstrs)

        cmd = f" return mcu.i2c.write16(0x{device:02X}, 0x{regaddr:04X}, {{{cmddstr}}})"

        # self.CmdSentSignal.emit(cmd)

        err, *_ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

        match err:
            case 0:
                return

            case 0x03:
                raise I2CNACKError(f"{device=:02X}")
            case 0x11:
                raise I2CBusError(f"{device=:02X}")
            case _:
                raise I2CError(f"{device=:02X}")

    # === Gary add on for JIGM3

    '''
    the mainly function will be I/O control and pattern gen
    PWM and SPI will be placed in low priority
    '''

    def g_ezcommand(self, command0):
        '''
        send string from V4 to get MCU work for cute g
        '''
        command0 = str(command0)
        if self.sim_mcu == 0:
            # object in simulation mode
            print(f'simulation mode for MCU, command is \n {command0}')
            pass
        else:
            print(f'real mode with command \n {command0}')
            self.ezCommand(command0)
            pass

        pass

    def i_o_config(self, port_sel0='PG', i_o_sel0='out', all_out0=1):
        '''
        setting up for the i_o input or output \n
        default output  \n
        port_sel need to use upper case  \n
        '''
        # MCU set 1 is output, 0 is input
        config_sel = 0xFFFF

        if i_o_sel0 == 'in':
            config_sel = 0x0
            pass

        if all_out0 != 1:
            # single setting request is send
            cmd_str = f'mcu.gpio.setdir({config_sel}, {self.i_o_port_num[port_sel0]})'
            self.g_ezcommand(cmd_str)

        else:
            # set all port to output
            self.g_ezcommand(f'mcu.gpio.setdir(0xFFFF, 0)')
            self.g_ezcommand(f'mcu.gpio.setdir(0xFFFF, 1)')
            self.g_ezcommand(f'mcu.gpio.setdir(0xFFFF, 2)')
            self.g_ezcommand(f'mcu.gpio.setdir(0xFFFF, 3)')

            # and send output to 0, so it can be aligned to default state
            self.g_ezcommand(f'mcu.gpio.setout(0x0, 0)')
            self.g_ezcommand(f'mcu.gpio.setout(0x0, 1)')
            self.g_ezcommand(f'mcu.gpio.setout(0x0, 2)')
            self.g_ezcommand(f'mcu.gpio.setout(0x0, 3)')

            pass

        pass

    def i_o_change(self, port0='PG', set_or_clr0=0, pin_num0=0):
        '''
        port0 default PG (0) \n
        set-1 or clr-0 \n
        pin_num is from 1-16 \n
        EX: PG1- PG16 \n
        '''
        port0 = str(port0)
        pin_num0 = str(pin_num0)

        port_cmd = self.i_o_port_num[port0]
        # put related IO state in cache
        i_o_state_tmp = self.i_o_state[port0]

        if set_or_clr0 == 1:
            # set the related pin
            i_o_cmd0 = self.i_o_pin_num_set[pin_num0]
            print(f'command0 is {i_o_cmd0}')
            # use or operator for the set
            i_o_cmd = i_o_state_tmp | i_o_cmd0
            print(f'final command0 is {i_o_cmd}')

        else:
            # clear the related pin
            i_o_cmd0 = self.i_o_pin_num_clr[pin_num0]
            print(f'command0 is {i_o_cmd0}')
            # use and operator for clear
            i_o_cmd = i_o_state_tmp & i_o_cmd0
            print(f'final command is {i_o_cmd}')

        cmd_str = f'mcu.gpio.setout({i_o_cmd}, {port_cmd})'
        # update the status of MCU
        self.i_o_state[port0] = i_o_cmd

        self.g_ezcommand(cmd_str)

        pass

    def pwm_ctrl(self, pwm_freq0=1000, pwm_duty0=10, pwm_port0=1, en_dis0=1):
        '''
        this function control the PWM output for the MCU
        PWM have 4 port PWM1-PWM4, default 1
        EX: mcu.pwm.setpwm(1000, 0 , 1) ; mcu.pwm.off()
        freq, duty, port, enalbe/disable
        '''

        if en_dis0 == 1:
            # enable PWM output
            cmd_str = f'mcu.pwm.setpwm({pwm_freq0}, {pwm_duty0} , {pwm_port0})'
            self.g_ezcommand(cmd_str)

            pass
        else:
            # disable the PWM output
            cmd_str = f'mcu.pwm.off()'
            self.g_ezcommand(cmd_str)

            pass

        pass

    def pattern_gen(self, pattern0=0, unit_time0=1, extra_function0=0):
        '''
        data format, ({data transition1, data transition2,...}, unit_time, extra_function)
        unit_time is in 'ns' \n
        format of data transition: (IO state,1-H,0-L)$(counter, unit_time) \n
        EX: mcu.pattern.setupX( {'9$2e4`401$1e4`409$1e4`400$1e4`0$5e4`'},
                    10000, 0 )
        '''
        '''
        pattern element break down
        9$2e4 => 9 is the status from PG16(MSB) to PG1(LSB) => PG1 and PG4 high
        9$2e4 => 2e4 is 2*10^4 * unit_time0 is the transition time point
        '' => is the beginning and end of the pattern element
        ` (~~) => is the separation of each transition of pattern
        '''

        if pattern0 == 0:
            # use default pattern, change all output to 0
            pattern0 = "'0$1e3`'"

        cmd_str = f"mcu.pattern.setupX( {{{pattern0}}},{unit_time0}, {extra_function0} )"
        self.g_ezcommand(cmd_str)

        pass

    def pattern_gen_full_str(self, cmd_str0=0):
        '''
        send the full string command generate from GPL,V4 pattern gen
        copy and paste~
        '''
        if cmd_str0 == 0:
            # use default pattern
            cmd_str = "mcu.pattern.setupX( {'0$1e3`'},10000, 0 )"

        cmd_str = cmd_str0
        self.g_ezcommand(cmd_str)

        pass

    def g_pulse_out(self, pulse0=1, low_duration_us=10):
        '''
        function to send pulse for SWIRE function
        '''

        pass

    '''
    add function to match MSP430
    '''

    def pulse_out(self, pulse_1, pulse_2):
        '''
        GPL MCU mapped with MSP430
        pulse need to be less then 255
        '''

        pass

    def pmic_mode(self, mode_index):
        '''
        (EN,SW) or (EN2, EN1) \n
        1:(0,0); 2:(0,1); 3:(1,0); 4:(1,1)
        default using PG16 as EN and PG15 as SW in JIGM3
        '''
        # mode index should be in 1-4
        if mode_index < 1 or mode_index > 4:
            mode_index = 1
            # turn off if error occur
        self.mode_set = mode_index

        if mode_index == 1:
            # shut_down
            self.i_o_change(self, port0='PG', set_or_clr0=0, pin_num0=16)
            self.i_o_change(self, port0='PG', set_or_clr0=0, pin_num0=15)
            pass
        elif mode_index == 2:

            pass


if __name__ == '__main__':
    # testing of GPL MCU connection
    '''
    only GPL MCU will be on the device, should be able to use the path[0]
    for default, when only one device is connected
    '''

    # GPL MCU initialization: this is the general initial method
    # for only one device at the same time
    # ====
    path = JIGM3.listdevices()
    g_mcu = JIGM3(devpath=path[0], sim_mcu0=1)
    # set all the IO to output
    g_mcu.i_o_config()
    # ====

    a = g_mcu.getversion()
    print(f'the MCU version is {a}')

    test_index = 3
    '''
    testing index settings

    '''

    if test_index == 0:
        # first to test the EZ command to config I/O of the GPL MCU

        x = 0x01
        y = 0x20

        print(x & y)
        print(x | y)
        print(x ^ y)

        print(~x)
        print(~y)

        pass

    elif test_index == 1:
        # send IO signal to MCU

        g_mcu.i_o_config(port_sel0='PG')

        # infinite toggle IO
        while (1):

            g_mcu.i_o_change(port0='IO', set_or_clr0=1, pin_num0=1)
            time.sleep(2)
            g_mcu.i_o_change(port0='IO', set_or_clr0=0, pin_num0=1)
            time.sleep(2)

            g_mcu.i_o_change(port0='IO', set_or_clr0=1, pin_num0=1)
            time.sleep(2)
            g_mcu.i_o_change(port0='IO', set_or_clr0=1, pin_num0=2)
            time.sleep(2)
            g_mcu.i_o_change(port0='IO', set_or_clr0=1, pin_num0=3)
            time.sleep(2)

            g_mcu.i_o_change(port0='IO', set_or_clr0=0, pin_num0=1)
            time.sleep(2)
            g_mcu.i_o_change(port0='IO', set_or_clr0=0, pin_num0=2)
            time.sleep(2)
            g_mcu.i_o_change(port0='IO', set_or_clr0=0, pin_num0=3)
            time.sleep(2)

            pass
        pass
    elif test_index == 2:
        '''
        testing for PWM control
        '''
        g_mcu.pwm_ctrl(pwm_freq0=45000, pwm_duty0=30, pwm_port0=1, en_dis0=1)

        g_mcu.pwm_ctrl(en_dis0=0)

        pass

    elif test_index == 3:
        '''
        testing for pattern gen
        { or } in f-string need to be {{ or }}
        '''
        pattern = "'0$1e2`20008$1e2`0$1e2`20008$1e2`0$6e2`'"
        full_str = "mcu.pattern.setupX( {'0$1e2`20008$1e2`0$1e2`20008$1e2`0$6e2`'},10000, 2 )"
        unit_time = 10000
        extra_function = 0
        g_mcu.pattern_gen(pattern0=pattern,
                          unit_time0=unit_time, extra_function0=2)
        time.sleep(5)
        g_mcu.pattern_gen_full_str(cmd_str0=full_str)

        pass

    elif test_index == 4:
        '''
        testing for pulse output function with programmable duration
        '''
