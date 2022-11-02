import pyvisa
import re

from .GInst import *
import time


class Power_BK9141(GInst):
    '''
    Class library from Geroge is channel based instrument, need to define by channel
    not a single instrument \n
    for parallel operation, need to setup by hand and use the channel 1 as control window
    '''

    def __init__(self, link, ch):
        super().__init__()

        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch

        self.chConvert = {'CH1': '0', 'CH2': '1', 'CH3': '2', 'ALL': ':ALL'}

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
            raise Exception(f'<>< Power_BK9141 ><> "9141" not fit in "{idn}"!')

    # @GInstSetMethod(unit = 'V')

    def setVoltage(self, voltage):
        """
        power.setVoltage(voltage) -> None
        ================================================================
        [power(channel) set Voltage]
        :param voltage:
        :return: None.
        """
        if voltage is None:
            return

        self.inst.write(f'INST {self.chConvert[self.ch]}')
        self.inst.write(f'VOLT {abs(voltage):3g}')

    # @GInstSetMethod(unit = 'A')

    def setCurrent(self, current):
        """
        power.setCurrent(current) -> None
        ================================================================
        [power(channel) set Current]
        :param current:
        :return: None.
        """
        if current is None:
            return

        self.inst.write(f'INST {self.chConvert[self.ch]}')
        self.inst.write(f'CURR {current:3g}')

    # @GInstOnMethod()

    def outputON(self):
        """
        power.outputON() -> None
        ================================================================
        [power(channel) output ON]
        :param None:
        :return: None.
        """
        self.inst.write(f'INST {self.chConvert[self.ch]}')
        self.inst.write(f'OUTP 1')

    # @GInstOffMethod()

    def outputOFF(self):
        """
        power.outputOFF() -> None
        ================================================================
        [power(channel) output OFF]
        :param None:
        :return: None.
        """
        self.inst.write(f'INST {self.chConvert[self.ch]}')
        self.inst.write(f'OUTP 0')

    # @GInstGetMethod(unit = 'V')

    def measureVoltage(self):
        """
        power.measureVoltage() -> Voltage
        ================================================================
        [power(channel) measure Voltage]
        :param None:
        :return: Voltage.
        """
        try:
            self.inst.write(f'INST {self.chConvert[self.ch]}')
            valstr = self.inst.query(f'MEAS:SCAL:VOLTage:DC?')
        except pyvisa.errors.VisaIOError:
            valstr = self.inst.query(f'MEAS:SCAL:VOLTage:DC?')
            self.inst.query(f'*CLS?')

        return float(re.search(r"[-+]?\d*\.\d+|\d+", valstr).group(0))

    # @GInstGetMethod(unit = 'A')

    def measureCurrent(self):
        """
        power.measureCurrent() -> Current
        ================================================================
        [power(channel) measure Current]
        :param None:
        :return: Current.
        """
        try:
            self.inst.write(f'INST {self.chConvert[self.ch]}')
            valstr = self.inst.query(f'MEAS:SCAL:CURR:DC?')
        except pyvisa.errors.VisaIOError:
            valstr = self.inst.query(f'MEAS:SCAL:CURR:DC?')
            self.inst.query(f'*CLS?')

        return float(re.search(r"[-+]?\d*\.\d+|\d+", valstr).group(0))
