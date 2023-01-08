
import logging
import json

# from PySide6.QtCore import *  # type: ignore
# from PySide6.QtGui import *  # type: ignore
# from PySide6.QtWidgets import *  # type: ignore

from Except import *

from PyUSBManagerv4 import opendev, closedev, luacmd, listdevs


class JIGM3():

    def __init__(self, devpath, CmdSentSignal=0, sim_mcu0=0, com_addr0=100):

        self.DevPath = devpath
        self.handle = opendev(devpath)

        self.CmdSentSignal = CmdSentSignal

        self.Version = self.getversion()

        logging.info(f"OpenJIGM3[{self.Version}]")

        # default mode setting will be simulation mode
        self.sim_mcu = sim_mcu0
        self.com_addr = com_addr0

        if self.com_addr == 100:
            # set MCU to simulation mode if the addr is set to 100
            self.sim_mcu = 0

        # because MCU will be separate with GPIB for implementation and test
        self.mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
        # meter channel indicator: 0: Vin, 1: AVDD, 2: OVDD, 3: OVSS, 4: VOP, 5: VON
        # array mpaaing for the relay control
        self.meter_ch_ctrl = 0
        self.array_rst = '00'

        self.pulse1 = 0
        self.pulse2 = 0
        self.mode_set = 4
        # mode sequence: 1-4: (EN, SW) = (0, 0),  (0, 1), (1, 0), (1, 1) => default normal

        # i2c register and data (single byte data); slave address is fixed in the MCU
        # here only support for the register and data change
        self.reg_i2c = ''
        self.data_i2c = ''

        #  UART command string
        self.uart_cmd_str = ''

        # different mode used in the operation
        self.mcu_mode_swire = 1
        self.mcu_mode_sw_en = 3
        self.mcu_mode_I2C = 4
        self.mcu_mode_8_bit_IO = 5
        self.mcu_mode_pat_gen_py = 6
        self.mcu_mode_pat_gen_encode = 7
        self.mcu_mode_pat_gen_direct = 8

        # MCU mapping for different mode control in 2553
        # MCU mapping for RA_GPIO control is mode 5
        # both mode 1 and 2 should defined as dual SWIRE, need to send two pulse command at one time
        # need to be build in the control sheet
        self.wait_time = 0.5
        self.wait_small = 0.2
        # waiting time duration, can be adjust with MCU C code

        pass

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
        '''
        re open
        '''
        if hasattr(self, "handle") and self.handle != None:
            closedev(self.handle)

        self.handle = opendev(self.DevPath)

    def close(self):
        if hasattr(self, "handle") and self.handle != None:
            closedev(self.handle)

    # -----------------------------------------------------------------------------------------------------------------
    def retry_and_refresh(func):
        '''
        decorator =>
        '''
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
        '''
        return things should be list from V4, check if error from JIG return (must be list if correct)
        '''
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
        option write i2c (usually for single byte, I2C didn' define register)
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
        read i2c (16 bit address)
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
        write i2c (16 bit resigeter)
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

    '''
    230106: new funcion added for matching to MSP430 tool
    '''

    def com_open(self):
        '''
        open COM port
        '''

        temp = 'sim mode'
        if self.sim_mcu == 0:
            temp = self.getversion()

        print(f'MCU open: {temp}')

        return str(temp)

    def com_close(self):
        # after the verification is finished, reset all the condition
        # to initial and turn off the communication port
        self.back_to_initial()
        print('the MCU will turn off')
        if self.sim_mcu == 1:
            # not doing anything
            pass
        else:
            print('the com port is turn off now(SIM)')

        pass

    def back_to_initial(self):
        # this sub program used to set all the MCU condition to initial
        # to change the initial setting, just modify the items from here

        # MCU will be in normal mode (EN, SW) = (1, 1) => 4
        self.pmic_mode(4)
        # the relay channel also reset to the default

        print('command accept to reset the MCU')
        pass

    def pulse_out(self, pulse_1=0, pulse_2=0, msp430=1):
        '''
        pulse need to be less than 255 for msp430, set msp430 to 0, break
        msp430 function limit
        '''
        if msp430 == 1:
            # pulse should be within 255 (8bit data limitation)
            if pulse_1 > 255:
                pulse_1 = 255
            if pulse_2 > 255:
                pulse_2 = 255
            pass

        self.pulse1 = pulse_1
        self.pulse2 = pulse_2
        if (self.pulse1 == 0) and (self.pulse2 == 0):
            print('pulse 0 0 is send, MCU no action')
            print('cute Grace!')
            pass
        if self.pulse1 != 0:
            self.g_pulse(self.pulse1)

        pass

    def g_pulse(self, pulse_count=0, i_o=0x02):
        '''
        to send the minimum duration low pulse toggle GPIO
        support different IO, but default SW(PG2)
        check on MCU command
        '''

        pass

    def pmic_mode(self, mode_index, msp430=1):
        '''
        for JIGM3, default: EN=PG1, SW=PG2
        (EN,SW) or (EN2, EN1) \n
        1:(0,0); 2:(0,1); 3:(1,0); 4:(1,1)
        '''
        # mode index should be in 1-4
        if mode_index < 1 or mode_index > 4:
            mode_index = 1
            # turn off if error occur
        self.mode_set = mode_index
        self.mcu_write('en_sw')
        pass

    def relay_ctrl(self, channel_index, msp430=1):
        '''
        index array: ['01', '02', '04', '08', '10', '20', '40', '80']\n
        from 0 to 7 \n
        MCU IO 2.0 - 2.7
        '''
        self.meter_ch_ctrl = int(channel_index)
        self.mcu_write('relay')
        pass

    def i2c_single_write(self, register_index, data_index, msp430=1):
        self.reg_i2c = register_index
        self.data_i2c = data_index
        self.mcu_write('i2c')

    # to update the implementation of other function
    # think about what is needed from the MCU operation
    # update later

    pass


if __name__ == '__main__':
    # testing of GPL MCU connection
    '''
    only GPL MCU will be on the device, should be able to use the path[0]
    for default, when only one device is connected
    '''

    path = JIGM3.listdevices()
    g_mcu = JIGM3(devpath=path[0])

    a = g_mcu.getversion()
    print(f'the MCU version is {a}')

    '''
    import the JIGM3, define g_mcu will be the GPL_tool MCU
    230106: also need to have GPIO toggling, pattern gen input? may not need, no used currently
    '''

    '''
    will need to consider the the MCU command already used in current object
    for the relay board, still plan for the old way of implementation
    control by IO from MCU, using GPIO, not to use GRELAY board
    since the relay switching now is implement by the IO and relay control
    '''
