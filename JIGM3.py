
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

        if set_or_clr0 == 1:
            # set the related pin
            i_o_cmd = self.i_o_pin_num_set[pin_num0]

        else:
            # clear the related pin
            i_o_cmd = self.i_o_pin_num_clr[pin_num0]

        cmd_str = f'mcu.gpio.setout({i_o_cmd}, {port_cmd})'

        self.g_ezcommand(cmd_str)

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
    # ====

    a = g_mcu.getversion()
    print(f'the MCU version is {a}')

    test_index = 1
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
