import pyvisa
import re

from .GInst import *


class Power_LPS505N(GInst):


    def __init__(self, link, ch) :
        super().__init__()

        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch

        self.chConvert = {'CH1':'1', 'CH2':'2', 'CH3':'3', 'ALL':':ALL'}

        try :
            self.inst = rm.open_resource(link)
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            self.inst.baud_rate = 38400

        except Exception as e:
            raise Exception(f'<>< Power_LPS505N ><> open 3-CH Power Fail {str(e)}!')


        idn = self.inst.query('*IDN?')

        if 'LPS-505N' not in idn and 'LPS505N' not in idn:
            raise Exception(f'<>< Power_LPS505N ><> "LPS-505N" not fit in "{idn}"!')


            

    @GInstSetMethod(unit = 'V')
    def setVoltage(self, voltage) :
        """
        power.setVoltage(voltage) -> None
        ================================================================
        [power(channel) set Voltage] 
        :param voltage: 
        :return: None.
        """
        if voltage is None :
            return

        self.inst.write(f'VOLT{self.chConvert[self.ch]} {abs(voltage):3g}')


    @GInstSetMethod(unit = 'A')
    def setCurrent(self, current) :
        """
        power.setCurrent(current) -> None
        ================================================================
        [power(channel) set Current] 
        :param current: 
        :return: None.
        """
        if current is None :
            return 
            
        self.inst.write(f'CURR{self.chConvert[self.ch]} {current:3g}')


    @GInstOnMethod()
    def outputON(self) :
        """
        power.outputON() -> None
        ================================================================
        [power(channel) output ON] 
        :param None: 
        :return: None.
        """
        self.inst.write(f'OUT{self.chConvert[self.ch]} 1')


    @GInstOffMethod()
    def outputOFF(self) :
        """
        power.outputOFF() -> None
        ================================================================
        [power(channel) output OFF] 
        :param None: 
        :return: None.
        """
        self.inst.write(f'OUT{self.chConvert[self.ch]} 0')


    @GInstGetMethod(unit = 'V')
    def measureVoltage(self):
        """
        power.measureVoltage() -> Voltage
        ================================================================
        [power(channel) measure Voltage] 
        :param None: 
        :return: Voltage.
        """
        valstr = self.inst.query(f'MEAS:VOLT{self.chConvert[self.ch]}?')
        return float(re.search(r"[-+]?\d*\.\d+|\d+", valstr).group(0))


    @GInstGetMethod(unit = 'A')
    def measureCurrent(self):
        """
        power.measureCurrent() -> Current
        ================================================================
        [power(channel) measure Current] 
        :param None: 
        :return: Current.
        """   
        valstr = self.inst.query(f'MEAS:CURR{self.chConvert[self.ch]}?')
        return float(re.search(r"[-+]?\d*\.\d+|\d+", valstr).group(0))

