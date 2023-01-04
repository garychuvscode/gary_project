
import logging
import json

# from PySide6.QtCore import *  # type: ignore
# from PySide6.QtGui import *  # type: ignore
# from PySide6.QtWidgets import *  # type: ignore

from Except import *

from PyUSBManagerv4 import opendev, closedev, luacmd, listdevs


class JIGM3():

    def __init__(self, devpath, CmdSentSignal=0):

        self.DevPath = devpath
        self.handle = opendev(devpath)

        self.CmdSentSignal = CmdSentSignal

        self.Version = self.getversion()

        logging.info(f"OpenJIGM3[{self.Version}]")

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
