import logging
import json
import time

# fmt: off
# from PySide6.QtCore import *  # type: ignore
# from PySide6.QtGui import *  # type: ignore
# from PySide6.QtWidgets import *  # type: ignore

from Except import *

from PyUSBManagerv4 import opendev, closedev, luacmd, listdevs


class JIGM3:
    def __init__(self, sim_mcu0=0, com_addr0=0, devpath=0, CmdSentSignal=0):
        self.DevPath = devpath
        try:
            self.handle = opendev(devpath)
            self.Version = self.getversion()

        except:
            self.handle = 1
            self.Version = "not open success"
            print("error or in simulation mode")

        self.CmdSentSignal = CmdSentSignal

        logging.info(f"OpenJIGM3[{self.Version}]")

        # === Gary add on for JIGM3
        """
        The add on for initial object is for default setting of JIGM3 or the generation
        of variable
        """

        self.sim_mcu = sim_mcu0
        self.com_addr = com_addr0
        # default config for PG1 and PG2 is EN and SW setting (default high)
        self.en_sw_pin = 1

        # port number assignment
        self.i_o_port_num = {"PG": "0", "IO": "1", "IOx": "2", "IOy": "3"}

        # IO number assignemnt
        self.i_o_pin_num_set = {
            "1": 0x1,
            "2": 0x2,
            "3": 0x4,
            "4": 0x8,
            "5": 0x10,
            "6": 0x20,
            "7": 0x40,
            "8": 0x80,
            "9": 0x100,
            "10": 0x200,
            "11": 0x400,
            "12": 0x800,
            "13": 0x1000,
            "14": 0x2000,
            "15": 0x4000,
            "16": 0x8000,
        }
        self.i_o_pin_num_clr = {
            "1": 0xFFFE,
            "2": 0xFFFD,
            "3": 0xFFFB,
            "4": 0xFFF7,
            "5": 0xFFEF,
            "6": 0xFFDF,
            "7": 0xFFBF,
            "8": 0xFF7F,
            "9": 0xFEFF,
            "10": 0xFDFF,
            "11": 0xFBFF,
            "12": 0xF7FF,
            "13": 0xEFFF,
            "14": 0xDFFF,
            "15": 0xBFFF,
            "16": 0x7FFF,
        }
        # binary number: 0b10101010

        # default state and state record of IO pin
        self.i_o_state = {"PG": 0x00, "IO": 0x00, "IOx": 0x00, "IOy": 0x00}

        """
        MSP430 related variable mapping
        """

        self.mcu_cmd_arry = ["01", "02", "04", "08", "10", "20", "40", "80"]

        # default PMIC mode is set to 1 (shut down mode)
        self.mode_set = 1

        # relay function enable or disable, not to use IO1-IO8 if relay mode enable
        self.relay0 = 0
        # define for temp relay channel in Vin calibration function
        self.meter_ch_ctrl = 0

    @staticmethod
    def listdevices():
        try:
            pathm4s = listdevs(
                (0x04D8),
                (0x003E),
                (
                    0x4D1E55B2,
                    0xF16F,
                    0x11CF,
                    0x88,
                    0xCB,
                    0x00,
                    0x11,
                    0x11,
                    0x00,
                    0x00,
                    0x30,
                ),
            )
            if pathm4s == []:
                path_null = ["no device or in simulation mode"]
                print("no device or in simulation mode, return null jig path")
                return path_null
            else:
                return pathm4s

        except Exception as e:
            logging.warning(e)
            return []

    # JIG-M4 style

    def getversion(self):
        # return self.ezCommand("return mcu._version:get()")[0]
        return self.g_ezcommand("return mcu._version:get()")[0]

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
            raise Exception("[Result Error] Result[0] is not integer for Error-number")

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
        """
        Gary Q:
        1. what is the return value? => if sending I/O pattern from pattern gen or I2C command?
        2. what is the difference between i2c i2c_op i2c16?
        A2: i2C16 is 16 bit data length => also address?; i2C should be the normal read an write with 8bit data and
        7 bit address
        Gary note:
        1. ezcommand can be found from the V4 tool "MCU", the string needed for related operation for JIGM3
        2. pattern gen, I/O toggle and I2C command supported
        """
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

        """
        return datas doen's contain register address
        230915: command example get from V4: UI-Link
        mcu.i2c.read(0x9C, 0xA0, 48)
        => read 48 byte from address 9C(8bit), start from register "A0"


        """

        cmd = f" return mcu.i2c.read(0x{device:02X}, 0x{regaddr:02X}, {len})"

        # self.CmdSentSignal.emit(cmd)

        # 230918: add the simulation mode setting for I2C related command
        # since I2C related items using the luacmd, not the ezcommand
        # need to add selection to prevent error in simulation mode

        if self.sim_mcu == 1 :
            # real mode, send the lua command and check for error

            err, datas, *_ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

            match err:
                case 0:
                    return datas

                case 0x03:
                    raise I2CNACKError(f"{device=:02X}")
                case 0x11:
                    raise I2CBusError(f"{device=:02X}")
                case _:
                    raise I2CError(f"{device=:02X}")

            pass
        else :
            # I2C related operation in simulation mode, just print the command
            print(f'send I2C command: "{cmd}" in sim mode')
            pass

    @retry_and_refresh
    def i2c_write(self, device, regaddr, datas):
        """
        write i2c
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param datas: list for bytes to write.
        :return: error id, datas or error string.
        """

        """
        device address use 8 bit address, regaddr (0x9C), datas (0xFF or 0x00)-list no need index
        """
        hexstrs = ["0x{:02X}".format(data) for data in datas]
        cmddstr = ", ".join(hexstrs)

        cmd = f" return mcu.i2c.write(0x{device:02X}, 0x{regaddr:02X}, {{{cmddstr}}})"

        # self.CmdSentSignal.emit(cmd)

        # 230918: add the simulation mode setting for I2C related command
        # since I2C related items using the luacmd, not the ezcommand
        # need to add selection to prevent error in simulation mode

        if self.sim_mcu == 1 :
            # real mode, send the lua command and check for error
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

            pass
        else :
            # I2C related operation in simulation mode, just print the command
            print(f'send I2C command: "{cmd}" in sim mode')
            pass

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

        """
        only device address, not register adress, difference between op and normal
        """

        cmd = f" return mcu.i2c.opread(0x{device:02X}, {len}, {1 if contAck else 0}, {1 if lastAck else 0})"

        # self.CmdSentSignal.emit(cmd)

        err, datas, *_ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

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

        err, datas, *_ = self.assertResult(json.loads(luacmd(self.handle, cmd)))

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

    """
    the mainly function will be I/O control and pattern gen
    PWM and SPI will be placed in low priority
    """

    def g_ezcommand(self, command0):
        """
        send string from V4 to get MCU work for cute g
        """
        command0 = str(command0)
        result = "default result"
        if self.sim_mcu == 0:
            # object in simulation mode
            print(f"simulation mode for MCU, command is \n {command0}")
            pass
        else:
            print(f"real mode with command \n {command0}")
            result = self.ezCommand(command0)
            # time.sleep(0.2)
            pass

        return result

    def i_o_config(self, port_sel0="PG", i_o_sel0="out", all_out0=1):
        """
        setting up for the i_o input or output \n
        default output  \n
        port_sel need to use upper case  \n
        """
        # MCU set 1 is output, 0 is input
        config_sel = 0xFFFF

        if i_o_sel0 == "in":
            config_sel = 0x0
            pass

        if all_out0 != 1:
            # single setting request is send
            cmd_str = f"mcu.gpio.setdir({config_sel}, {self.i_o_port_num[port_sel0]})"
            self.g_ezcommand(cmd_str)

        else:
            # set all port to output
            self.g_ezcommand(f"mcu.gpio.setdir(0xFFFF, 0)")
            self.g_ezcommand(f"mcu.gpio.setdir(0xFFFF, 1)")
            self.g_ezcommand(f"mcu.gpio.setdir(0xFFFF, 2)")
            self.g_ezcommand(f"mcu.gpio.setdir(0xFFFF, 3)")

            # and send output to 0, so it can be aligned to default state
            if self.en_sw_pin == 1:
                # self.g_ezcommand(f"mcu.gpio.setout(0x3, 0)")
                self.i_o_change(port0="PG", set_or_clr0=1, pin_num0=1)
                self.i_o_change(port0="PG", set_or_clr0=1, pin_num0=2)
                for i in range(3, 16, 1):
                    self.i_o_change(port0="PG", set_or_clr0=1, pin_num0=i)

                pass
            else:
                for i in range(1, 16, 1):
                    self.i_o_change(port0="PG", set_or_clr0=1, pin_num0=i)

                # self.g_ezcommand(f"mcu.gpio.setout(0x0, 0)")
            self.g_ezcommand(f"mcu.gpio.setout(0x0, 1)")
            self.g_ezcommand(f"mcu.gpio.setout(0x0, 2)")
            self.g_ezcommand(f"mcu.gpio.setout(0x0, 3)")

            pass

        pass

    def i_o_change(self, port0="PG", set_or_clr0=0, pin_num0=0):
        """
        port0 default PG (0) \n
        set-1 or clr-0 \n
        pin_num is from 1-16 \n
        EX: PG1- PG16 \n
        """
        port0 = str(port0)
        pin_num0 = str(pin_num0)

        port_cmd = self.i_o_port_num[port0]
        # put related IO state in cache
        i_o_state_tmp = self.i_o_state[port0]

        if set_or_clr0 == 1:
            # set the related pin
            i_o_cmd0 = self.i_o_pin_num_set[pin_num0]
            print(f"command0 is {i_o_cmd0}")
            # use or operator for the set
            i_o_cmd = i_o_state_tmp | i_o_cmd0
            print(f"final command0 is {i_o_cmd}")

        else:
            # clear the related pin
            i_o_cmd0 = self.i_o_pin_num_clr[pin_num0]
            print(f"command0 is {i_o_cmd0}")
            # use and operator for clear
            i_o_cmd = i_o_state_tmp & i_o_cmd0
            print(f"final command is {i_o_cmd}")

        if port0 == "PG":
            # when use PG pin as IO, need to drive with pattern gen function
            cmd_str = f"mcu.pattern.setupX( {{'{i_o_cmd}$1e5`'}}, 10000, 0 )"
            pass
        else:
            cmd_str = f"mcu.gpio.setout({i_o_cmd}, {port_cmd})"

        self.g_ezcommand(cmd_str)

        # update the status of MCU
        self.i_o_state[port0] = i_o_cmd

        pass

    def pwm_ctrl(self, pwm_freq0=1000, pwm_duty0=10, pwm_port0=1, en_dis0=1):
        """
        this function control the PWM output for the MCU
        PWM have 4 port PWM1-PWM4, default 1
        EX: mcu.pwm.setpwm(1000, 0 , 1) ; mcu.pwm.off()
        freq, duty, port, enalbe/disable
        """

        if en_dis0 == 1:
            # enable PWM output
            cmd_str = f"mcu.pwm.setpwm({pwm_freq0}, {pwm_duty0} , {pwm_port0})"
            self.g_ezcommand(cmd_str)

            pass
        else:
            # disable the PWM output
            cmd_str = f"mcu.pwm.off()"
            self.g_ezcommand(cmd_str)

            pass

        pass

    def pattern_gen(self, pattern0=0, unit_time_ns0=1, extra_function0=0):
        """
        data format, ({data transition1, data transition2,...}, unit_time, extra_function)
        unit_time is in 'ns' \n
        format of data transition: (IO state,1-H,0-L)$(counter, unit_time) \n
        EX: mcu.pattern.setupX( {'9$2e4`401$1e4`409$1e4`400$1e4`0$5e4`'},
                    10000, 0 )
        """
        """
        pattern element break down
        9$2e4 => 9 is the status from PG16(MSB) to PG1(LSB) => PG1 and PG4 high
        9$2e4 => 2e4 is 2*10^4 * unit_time_ns0 is the transition time point
        '' => is the beginning and end of the pattern element
        ` (~~) => is the separation of each transition of pattern
        """

        if pattern0 == 0:
            # use default pattern, change all output to 0
            pattern0 = "'0$1e3`'"

        cmd_str = (
            f"mcu.pattern.setupX( {{{pattern0}}},{unit_time_ns0}, {extra_function0} )"
        )
        self.g_ezcommand(cmd_str)

        pass

    def pattern_gen_full_str(self, cmd_str0=0):
        """
        send the full string command generate from GPL,V4 pattern gen
        copy and paste~
        """
        if cmd_str0 == 0:
            # use default pattern
            cmd_str = "mcu.pattern.setupX( {'0$1e3`'},10000, 0 )"

        cmd_str = cmd_str0
        self.g_ezcommand(cmd_str)

        pass

    def g_pulse_out(self, pulse0=1, duration_ns=1000, en_sw="SW"):
        """
        function to send pulse for SWIRE function
        output will be 10*duration !! \n
        PG1(EN) and PG2(SW) as output
        """

        """
        pattern gen explanation
        by using PG1(EN) and PG2(SW) as output
        toggle SW => '3$10`1$10`3$10`1$10`3$10`1$10`3$10`1$10`3$20`'
        (head and end => keep the pin in logic H: '3$10  `3$10`')
        format: each pulse add one cycle `1$10`3$10
        scalling is scalling to us (1000 ns, and ns is the duration unit)
        """
        cmd_str = "'3$10"
        cmd_str_end = "`'"

        # decide which pin to toggle
        if en_sw == "SW":
            # toggle SW
            single_cell = "`1$10`3$10"
            pass
        else:
            # toggle EN
            single_cell = "`2$10`3$10"

        # cmd_str_end = "`3$10`'"

        cmd_str = "'3$10"
        x = 0
        while x < pulse0:
            cmd_str = cmd_str + single_cell
            x = x + 1
            pass
        # after finished the pulse count, add the final element
        cmd_str = cmd_str + cmd_str_end
        self.pattern_gen(pattern0=cmd_str, unit_time_ns0=duration_ns)

        print("pulseV1 finished, Grace buy family mart for lunch")

        pass

    """
    add function to match MSP430
    """

    def pulse_out(self, pulse_1=1, pulse_2=1):
        """
        GPL MCU mapped with MSP430
        pulse need to be less then 255
        can choose pulse at PG2(EN), default is at PG1(SW)

        g_MCU LSB is 1 ns, set in duration to 10 us
        """
        """
        pattern gen explanation
        by using PG1(EN) and PG2(SW) as output
        toggle SW => '3$10`1$10`3$10`1$10`3$10`1$10`3$10`1$10`3$20`'
        (head and end => keep the pin in logic H: '3$10  `3$10`')
        format: each pulse add one cycle `1$10`3$10
        """

        # since this function is used to mapped with MSP430, fixed the extra
        # parameter of pattern gen
        duration_ns = 1000
        en_sw = "SW"

        cmd_str = "'3$10"
        cmd_str_end = "`'"

        if en_sw == "SW":
            # toggle SW
            single_cell = "`1$10`3$10"
            pass
        else:
            # toggle EN
            single_cell = "`2$10`3$10"

        # cmd_str_end = "`3$10`'"

        cmd_str = "'3$10"
        x = 0
        while x < pulse_1:
            cmd_str = cmd_str + single_cell
            x = x + 1
            pass
        # after finished the pulse count, add the final element
        cmd_str = cmd_str + cmd_str_end
        self.pattern_gen(pattern0=cmd_str, unit_time_ns0=duration_ns)

        print("pulse1 finished")

        # delay 50ms between two pulse
        time.sleep(0.05)

        cmd_str = "'3$10"
        x = 0
        while x < pulse_2:
            cmd_str = cmd_str + single_cell
            x = x + 1
            pass
        # after finished the pulse count, add the final element
        cmd_str = cmd_str + cmd_str_end
        self.pattern_gen(pattern0=cmd_str, unit_time_ns0=duration_ns)

        print("pulse2 finished")

        pass

    def pmic_mode(self, mode_index=1, port_optional0="PG", pin_EN0=1, pin_SW0=2):
        """
        (EN,SW) or (EN2, EN1) \n
        1:(0,0); 2:(0,1); 3:(1,0); 4:(1,1)
        default using PG1 as EN and PG2 as SW in JIGM3

        port_optional0, pin_EN0, pin_SW0 refer to the IO settings
        """

        # mode index should be in 1-4
        if mode_index < 1 or mode_index > 4:
            mode_index = 1
            # turn off if error occur
        self.mode_set = mode_index

        if mode_index == 1:
            # shut_down
            # 230712 since there need to change IO status at the same time, change to pattern gen operation
            self.pattern_gen(pattern0="'0$1e5`'")
            # self.i_o_change(port0=port_optional0, set_or_clr0=0, pin_num0=pin_EN0)
            # self.i_o_change(port0=port_optional0, set_or_clr0=0, pin_num0=pin_SW0)
            pass
        elif mode_index == 2:
            # only SW on
            # 230712 since there need to change IO status at the same time, change to pattern gen operation
            self.pattern_gen(pattern0="'2$1e5`'")
            # self.i_o_change(port0=port_optional0, set_or_clr0=0, pin_num0=pin_EN0)
            # self.i_o_change(port0=port_optional0, set_or_clr0=1, pin_num0=pin_SW0)
            pass
        elif mode_index == 3:
            # only EN on (AOD mode for PMIC)
            # 230712 since there need to change IO status at the same time, change to pattern gen operation
            self.pattern_gen(pattern0="'1$1e5`'")
            # self.i_o_change(port0=port_optional0, set_or_clr0=1, pin_num0=pin_EN0)
            # self.i_o_change(port0=port_optional0, set_or_clr0=0, pin_num0=pin_SW0)
            pass
        elif mode_index == 4:
            # both EN and SW are on (normal mode of PMIC)
            # 230712 since there need to change IO status at the same time, change to pattern gen operation
            self.pattern_gen(pattern0="'3$1e5`'")
            # self.i_o_change(port0=port_optional0, set_or_clr0=1, pin_num0=pin_EN0)
            # self.i_o_change(port0=port_optional0, set_or_clr0=1, pin_num0=pin_SW0)
            pass

        print(f'pmic moe set to {mode_index}\n')

        pass

    def com_open(self):
        """
        only reserve the function for MSP430 function mapping, prevent function call error
        """
        if self.sim_mcu == 1:
            # re-run the process of open JIGM3
            path_list = self.listdevices()
            path = path_list[0]

            self.DevPath = path
            self.handle = opendev(path)

            self.CmdSentSignal = 9

            self.Version = self.getversion()

            logging.info(f"OpenJIGM3[{self.Version}]")

            print(f"the device get from com_open is{self.Version}")

            # also add IO config into the com open function
            self.i_o_config()

        else:
            print("open in simulation mode")

        # update the path for JIGM3
        print("Grace just take in charge of the efficiency environment")

        # 230516
        # run the IO config after COM open, son no need extra function at main

        # set all the IO to output
        self.i_o_config()
        # default setting for EN and SWIRE to high
        self.pmic_mode(4)

        pass

    def com_close(self):
        """
        only reserve the function for MSP430 function mapping, prevent function call error
        """
        print("Grace is cursing the layout team from MAtek XD")

        self.close()
        pass

    def back_to_initial(self):
        """
        make the MCU back to initial state
        default normal mode, EN=SW=H
        """
        # MCU will be in normal mode (EN, SW) = (1, 1) => 4
        self.pmic_mode(mode_index=4)
        # the relay channel also reset to the default

        print("command accept to reset the MCU, Grace is late to office today XD")

        pass

    def relay_ctrl(self, channel_index=0, relay_mode0=1, t_dly_s=0.05):
        """
        MSP relay control \n
        use IO1(index0)-IO8(index7) as the related \n
        need to set relay mode to 1 at the first call of relay control \n
        this function only control 8 channel of relay, just MSP430 \n
        """
        # to use IO port, port command fixed to 1 in this function
        port_cmd = 1
        # set the control index to 1 => means the IO1-IO8 is set to
        # relay control

        # 230628 add the channel record for power supply Vin calibration
        self.meter_ch_ctrl = int(channel_index)

        if relay_mode0 == 1:
            # set relay mode to 1 and disable other GIPO control
            self.relay0 = 1
        elif relay_mode0 == 0:
            # turn off relay mode and return the control of GPIO
            self.relay0 = 0
        else:
            # not doing anything
            pass

        if self.relay0 == 1:
            # re-load other GPIO status setting from IO9-IO16
            # clear the PG1-PG8
            self.relay_state = self.i_o_state["IO"] & 0xFF00

            time.sleep(t_dly_s)
            # since relay control need to be causion for short between relays
            # give all turn off first
            cmd_str = f"mcu.gpio.setout({self.relay_state}, {port_cmd})"
            self.g_ezcommand(cmd_str)
            time.sleep(t_dly_s)

            cmd_str = (
                f"mcu.gpio.setout({self.i_o_pin_num_set[str(channel_index + 1)]}, {port_cmd})"
            )
            self.g_ezcommand(cmd_str)
            time.sleep(t_dly_s)
            print("Grace is looking for GPIB control tool")

        pass

    def glitch_test_H(
        self, start_us0=10, step_us0=10, stop_us0=500, pin_num0=0, port0="PG"
    ):
        """
        this function is planned to do the glitch testing of IO pin with
        but pattern generator can only use ns as unit, better use us
        as the testing step

        set-1 or clr-0 \n
        pin_num is from 1-16 \n
        EX: PG1- PG16 \n
        follow i_o_change

        for H_version (high pulse) select pin and the other state will be 0 (all pin0)
        """

        # config the pin for deglitch

        i_o_cmd0 = self.i_o_pin_num_set[pin_num0]

        c_glitch = 0
        x_glitch = int(stop_us0 / start_us0)
        while c_glitch < x_glitch:
            x_glitch = x_glitch + 1
            pass

        pass

    def g_pulse_out_V2(self, pulse0=1, duration_ns=1000, en_sw="SW", count0=1):
        """
        function to send pulse for SWIRE function
        output will be count0*duration !! \n
        by using PG1(EN) and PG2(SW) as output
        """

        """
        pattern gen explanation
        by using PG1(EN) and PG2(SW) as output
        toggle SW => '3$10`1$10`3$10`1$10`3$10`1$10`3$10`1$10`3$20`'
        (head and end => keep the pin in logic H: '3$10  `3$10`')
        format: each pulse add one cycle `1$10`3$10
        scalling is scalling to us (1000 ns, and ns is the duration unit)
        """
        cmd_str = "'3$1"
        cmd_str_end = "`'"
        count0 = int(count0)

        # decide which pin to toggle
        if en_sw == "SW":
            # toggle SW (when EN is high)
            single_cell = f"`1${count0}`3${count0}"
            pass
        else:
            # toggle EN (when SW is high)
            single_cell = f"`2${count0}`3${count0}"

        # cmd_str_end = "`3$10`'"

        cmd_str = "'3$1"
        x = 0
        while x < pulse0:
            cmd_str = cmd_str + single_cell
            x = x + 1
            pass
        # after finished the pulse count, add the final element
        cmd_str = cmd_str + cmd_str_end
        self.pattern_gen(pattern0=cmd_str, unit_time_ns0=duration_ns)

        print("g_pulse_V2 finished, Grace doesn't return the tweezers")

        pass


if __name__ == "__main__":
    # testing of GPL MCU connection
    """
    only GPL MCU will be on the device, should be able to use the path[0]
    for default, when only one device is connected
    """

    # GPL MCU initialization: this is the general initial method
    # for only one device at the same time
    # ====
    path = JIGM3.listdevices()
    g_mcu = JIGM3(devpath=path[0], sim_mcu0=1)
    # set simulation mode or normal mode
    g_mcu.sim_mcu = 1
    g_mcu.com_open()
    # set all the IO to output

    # 230627 since this function been put into the com_open, no need this command here
    # g_mcu.i_o_config()

    # ====

    a = g_mcu.getversion()
    print(f"the MCU version is {a}")

    test_index = 7
    """
    testing index settings
    1 => IO control
    2 => PWM control
    3 => pattern gen
    4 => pulse output for deglitch function or SWIRE
    5 => IO toggle for relay function of MSP(IO1-IO8)
    6 => PMIC mode
    7 => I2C read and write testing

    """

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

        # since all out is set to 1, no need specific port_sel, then
        # system set all pin to output automatically
        g_mcu.i_o_config(port_sel0="PG")

        # infinite toggle IO
        while 1:
            g_mcu.i_o_change(port0="IO", set_or_clr0=1, pin_num0=1)
            time.sleep(2)
            g_mcu.i_o_change(port0="IO", set_or_clr0=0, pin_num0=1)
            time.sleep(2)

            g_mcu.i_o_change(port0="IO", set_or_clr0=1, pin_num0=1)
            time.sleep(2)
            g_mcu.i_o_change(port0="IO", set_or_clr0=1, pin_num0=2)
            time.sleep(2)
            g_mcu.i_o_change(port0="IO", set_or_clr0=1, pin_num0=3)
            time.sleep(2)

            g_mcu.i_o_change(port0="IO", set_or_clr0=0, pin_num0=1)
            time.sleep(2)
            g_mcu.i_o_change(port0="IO", set_or_clr0=0, pin_num0=2)
            time.sleep(2)
            g_mcu.i_o_change(port0="IO", set_or_clr0=0, pin_num0=3)
            time.sleep(2)

            g_mcu.i_o_change(port0="PG", set_or_clr0=1, pin_num0=1)
            time.sleep(2)
            g_mcu.i_o_change(port0="PG", set_or_clr0=1, pin_num0=2)
            time.sleep(2)
            g_mcu.i_o_change(port0="PG", set_or_clr0=0, pin_num0=1)
            time.sleep(2)

            pass
        pass
    elif test_index == 2:
        """
        testing for PWM control
        """
        # turn on
        g_mcu.pwm_ctrl(pwm_freq0=45000, pwm_duty0=30, pwm_port0=1, en_dis0=1)

        # turn off
        g_mcu.pwm_ctrl(en_dis0=0)

        pass

    elif test_index == 3:
        """
        testing for pattern gen
        { or } in f-string need to be {{ or }}
        """
        pattern = "'0$1e2`20008$1e2`0$1e2`20008$1e2`0$6e2`'"
        full_str = (
            "mcu.pattern.setupX( {'0$1e2`20008$1e2`0$1e2`20008$1e2`0$6e2`'},10000, 2 )"
        )
        unit_time = 10000
        extra_function = 0
        g_mcu.pattern_gen(pattern0=pattern, unit_time_ns0=unit_time, extra_function0=2)
        time.sleep(5)
        g_mcu.pattern_gen_full_str(cmd_str0=full_str)

        pass

    elif test_index == 4:
        """
        testing for pulse output function with programmable duration
        """

        g_mcu.pulse_out(pulse_1=10, pulse_2=10)

        g_mcu.g_pulse_out(pulse0=5, duration_ns=1000, en_sw="SW")

        g_mcu.g_pulse_out(pulse0=5, duration_ns=1000, en_sw="EN")

        g_mcu.g_pulse_out_V2(pulse0=2, duration_ns=1000, en_sw="SW", count0=3)

        g_mcu.g_pulse_out_V2(pulse0=3, duration_ns=1000, en_sw="SW", count0=10)

        g_mcu.g_pulse_out_V2(pulse0=1, duration_ns=500, en_sw="SW", count0=10)

        g_mcu.com_close()

        pass

    elif test_index == 5 :
        # add the testing functino for relay, should be same with IO pin
        g_mcu.relay_ctrl(channel_index=0)
        time.sleep(0.1)
        g_mcu.relay_ctrl(channel_index=1)
        time.sleep(0.1)
        g_mcu.relay_ctrl(channel_index=2)
        time.sleep(0.1)
        g_mcu.relay_ctrl(channel_index=3)
        time.sleep(0.1)
        g_mcu.relay_ctrl(channel_index=4)
        time.sleep(0.1)
        g_mcu.relay_ctrl(channel_index=5)
        time.sleep(0.1)
        g_mcu.relay_ctrl(channel_index=6)
        time.sleep(0.1)
        g_mcu.relay_ctrl(channel_index=7)

        pass

    elif test_index == 6 :
        #  PMIC mode toggling testing
        while 1 :
            g_mcu.pmic_mode(mode_index=1)
            time.sleep(1)
            g_mcu.pmic_mode(mode_index=2)
            time.sleep(1)
            g_mcu.pmic_mode(mode_index=1)
            time.sleep(1)
            g_mcu.pmic_mode(mode_index=2)
            time.sleep(1)
            pass

    elif test_index == 7 :
        '''
        this setcition is the testing function of the I2C
        since there are verfication item need to use I2C change and scan
        MSP430 don't have prover I2C function yet, JIGM3 is the best choice

        use 374 as the testing circuit for I2C module

        '''

        # 0x9C in 8 bit address, means the 0x4E in the 7 bit general addressig of I2C
        addr = 0x9C

        # data_String = [ 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x10 ]
        # data_string = [ '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']

        # data should be loaded directly from the excel in string format

        # device0 = '9C'
        # register0 = 'A2'

        # the datas input need to be list, see if able to change to hex input
        data0 = [255, 128, 0, 30]
        # data1 = '8A'

        device0 = 0x9C
        register0 = 0xA2
        # data0 = 0x0A
        # data1 = 0x8A

        while 1 :
            data_f = g_mcu.i2c_read(device=device0, regaddr=register0, len=5)
            print(f'the readed data is: {data_f}.')
            time.sleep(3)

            g_mcu.i2c_write(device=device0, regaddr=register0, datas=data0)
            print(f'the wrote data is: {data0} in register {register0}.')
            time.sleep(3)

            g_mcu.i2c_read(device=device0, regaddr=register0, len=5)
            print(f'the readed data is: {data_f}.')
            time.sleep(3)

            print(f'the wrote data is: {data0} in register {register0}.')
            print(f'the readed data is: {data_f}.')
            time.sleep(3)




            pass

        pass
