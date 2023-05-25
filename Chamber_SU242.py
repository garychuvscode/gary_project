"""
since this library is from Geroge, need to modify before mapped to the testing system
1. need to have the fake GInst in the related folder
2. need to cancel all the deccorator
3. check how to build up the interface to connect two library
"""

import pyvisa
import re

from .GInst import *

# turn off the formatter
# fmt: off

class Chamber_SU242(GInst):

    def __init__(self, link, ch):
        super().__init__()

        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch

        try:
            self.inst = rm.open_resource(link)
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            self.inst.baud_rate = 38400
        except Exception as e:
            raise Exception(
                f'<>< Chamber_SU242 ><> open Chamber Fail {str(e)}!')

        idn = self.inst.query('*IDN?')

        if 'NA:CMD_ERR' not in idn:
            raise Exception(
                f'<>< Chamber_SU242 ><> Espec SU242 chamber init fail."{idn}"!')

    # -----------------------------------------------------------------------------------------------------------------
    # --  EXPORT (thread-safe) ----------------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------
    # @GInstOnMethod()
    # @GInstSetMethod(unit = 'C')

    def setTemp(self, temperature):
        """
        chamber.setTemp(temperature) -> None
        ================================================================
        [chamber setTemp]
        :param temperature:
        :return: None.
        """
        self.inst.write(f'MODE, CONSTANT')
        self.inst.write(f'TEMP, S{temperature:.1f}')

    @GInstGetMethod(unit='C')
    def getTemp(self):
        """
        chamber.getTemp() -> temperature
        ================================================================
        [chamber getTemp]
        :param :
        :return: temperature.
        """
        valstr = self.inst.query(f'TEMP?', 0.5)
        return float(re.search(r"[-+]?\d*\.\d+|\d+", valstr).group(0))

    @GInstOffMethod()
    def stop(self):
        """
        chamber.stop() -> None
        ================================================================
        [chamber stop]
        :param None:
        :return: None.
        """
        self.inst.write(f'MODE, STANDBY')

    # @GInstSetMethod(unit = 'C', value2=True)

    def waitTemp(self, temperature, diff=1.9):
        """
        chamber.waitTemp(temperature, diff = 1.9) -> None
        ================================================================
        [chamber stop]
        :param temperature:
        :param diff(option): tolerance
        :return: None.
        """
        from NAGlib.Gsys.Gsys import Gsys

        self.setTemp(temperature)
        Gsys.sleep(1)

        while True:
            if abs(self.getTemp() - temperature) < diff:
                return True

            Gsys.sleep(30)
