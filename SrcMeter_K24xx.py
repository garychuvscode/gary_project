'''
since this library is from Geroge, need to modify before mapped to the testing system
1. need to have the fake GInst in the related folder
2. need to cancel all the deccorator
3. check how to build up the interface to connect two library
'''

import pyvisa
import re

from .GInst import *


class SrcMeter_K24xx(GInst):


    def __init__(self, link, ch = 0, type = 'Loader') :
        super().__init__()

        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch
        self._type = type
        self._configured = False

        self.chConvert = {'CH1':'1', 'CH2':'2', 'CH3':'3', 'ALL':':ALL'}

        try :
            self.inst = rm.open_resource(link)
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            self.inst.baud_rate = 38400
        except Exception as e:
            raise Exception(f'<>< SrcMeter_K24XX ><> open FuncGen Fail {str(e)}!')

        self.inst.write(f'*RST')
        idn = self.inst.query('*IDN?')

        #source meter model#
        if '2400' in idn:
            self.inst_type = "2400"
        elif '2440' in idn:
            self.inst_type = "2440"
        elif '2450' in idn:
            self.inst_type = "2450"
        else:
            raise Exception(f'<>< SrcMeter_K24XX ><> 2400/2440 not fit in "{idn}" !')

        #if '2400' not in idn and '2440' not in idn:
        #    raise IOError('2400/2440 not fit in "{}"!'.format(idn))


    def autoConfigure(self):
        if self._configured is True :
            return

        if self._type == 'Loader' :
            self.inst.write(f':SOUR:FUNC CURR')
            self.inst.write(f':SOUR:CURR:MODE FIXED')
            self.inst.write(f':SOUR:CURR:RANG:AUTO ON')

            self.inst.write(f':SENS:FUNC "VOLT"')
            self.inst.write(f':SENS:VOLT:RANG:AUTO ON')

        if self._type == 'Power' :
            self.inst.write(f':SOUR:FUNC VOLT')
            self.inst.write(f':SOUR:VOLT:MODE FIXED')
            self.inst.write(f':SOUR:VOLT:RANG:AUTO ON')

            self.inst.write(f':SENS:FUNC "CURR"')
            self.inst.write(f':SENS:CURR:RANG:AUTO ON')

        self._configured = True


    # @GInstSetMethod(unit = 'V')  #decorate used to update UI
    def configureVoltageSource(self, voltage, current) :
        """
        srcmeter.configureVoltageSource(voltage, current) -> None
        ================================================================
        [srcmeter(channel) configure as VoltageSource ]
        :param voltage: Voltage source
        :param current: Current Limit
        :return: None.
        """
        Vrange = "AUTO ON"
        Irange = "AUTO ON"
        #choose Vrange
        if '2400' in self.inst_type:
            if voltage>20:
                Vrange = "200"
            elif voltage>2:
                Vrange = "20"
            elif voltage>1:
                Vrange = "10"
            elif voltage>0.2:
                Vrange = "2"
            elif voltage>0.1:
                Vrange = "0.2"
            else:
                Vrange = "AUTO ON"
        elif '2440' in self.inst_type:
            if voltage>20:
                Vrange = "40"
            elif voltage>2:
                Vrange = "20"
            elif voltage>1:
                Vrange = "10"
            elif voltage>0.2:
                Vrange = "2"
            elif voltage>0.1:
                Vrange = "0.2"
            else:
                Vrange = "AUTO ON"
        elif '2450' in self.inst_type:
            if voltage>20:
                Vrange = "200"
            elif voltage>2:
                Vrange = "20"
            elif voltage>1:
                Vrange = "10"
            elif voltage>0.2:
                Vrange = "2"
            elif voltage>0.1:
                Vrange = "0.2"
            else:
                Vrange = "AUTO ON"
        else:
            Vrange = "AUTO ON"

        #choose Irange
        if '2400' in self.inst_type:
            if current<100E-9:
                Irange = "100E-9"
            elif current<1000E-9:
                Irange = "1E-6"
            elif current<10E-6:
                Irange = "10E-6"
            elif current<100E-6:
                Irange = "100E-6"
            elif current<1000E-6:
                Irange = "1E-3"
            elif current<10E-3:
                Irange = "10E-3"
            elif current<100E-3:
                Irange = "100E-3"
            elif current<1000E-3:
                Irange = "1"
            else:
                Irange = "AUTO ON"
        elif '2440' in self.inst_type:
            if current<100E-9:
                Irange = "100E-9"
            elif current<1000E-9:
                Irange = "1E-6"
            elif current<10E-6:
                Irange = "10E-6"
            elif current<100E-6:
                Irange = "100E-6"
            elif current<1000E-6:
                Irange = "1E-3"
            elif current<10E-3:
                Irange = "10E-3"
            elif current<100E-3:
                Irange = "100E-3"
            elif current<1000E-3:
                Irange = "1"
            else:
                Irange = "AUTO ON"
        elif '2450' in self.inst_type:
            if current<100E-9:
                Irange = "100E-9"
            elif current<1000E-9:
                Irange = "1E-6"
            elif current<10E-6:
                Irange = "10E-6"
            elif current<100E-6:
                Irange = "100E-6"
            elif current<1000E-6:
                Irange = "1E-3"
            elif current<10E-3:
                Irange = "10E-3"
            elif current<100E-3:
                Irange = "100E-3"
            elif current<1000E-3:
                Irange = "1"
            else:
                Irange = "AUTO ON"
        else:
            Irange = "AUTO ON"

        #self.inst.write(f'VOLT{self.chConvert[self.ch]} {abs(voltage):3g}')
        self.inst.write(f':SENS:CURR:RANG:AUTO ON')
        self.inst.write(f':SOUR:FUNC VOLT')
        self.inst.write(f':SOUR:VOLT:MODE FIXED')

        if 'AUTO ON' in Vrange:
            self.inst.write(f':SOUR:VOLT:RANG:AUTO ON')
            #print(f':SOUR:VOLT:RANG:AUTO ON')
        else:
            self.inst.write(f':SOUR:VOLT:RANG {Vrange}')
            #print(f':SOUR:VOLT:RANG {Vrange}')
        self.inst.write(f':SOUR:VOLT:LEV {voltage}')

        self.inst.write(f':SENS:FUNC "CURR"')
        self.inst.write(f':SENS:CURR:PROT {current}')
        self.inst.write(f':SENS:CURR:RANG:AUTO ON')

        self._configured = True
        """
        if 'AUTO ON' in Irange:
            self.inst.write(f':SENS:CURR:RANG:AUTO ON')
            print(f':SENS:CURR:RANG:AUTO ON')
        else:
            self.inst.write(f':SENS:CURR:RANG {Irange}')
            print(f':SENS:CURR:RANG {Irange}')
        """




    # @GInstSetMethod(unit = 'A')
    def configureCurrentSource(self, voltage, current) :
        """
        srcmeter.configureCurrentSource(voltage, current) -> None
        ================================================================
        [srcmeter(channel) configure as CurrentSource]
        :param voltage: Voltage Limit
        :param current: Source Currnet
        :return: None.
        """
        Vrange = "AUTO ON"
        Irange = "AUTO ON"
        #choose Vrange
        if '2400' in self.inst_type:
            if voltage>200:
                Vrange = "AUTO ON"
            elif voltage<210E-3:
                Vrange = "AUTO ON"
            else:
                Vrange = "{:.2f}".format(voltage)
        elif '2440' in self.inst_type:
            if voltage>42:
                Vrange = "AUTO ON"
            elif voltage<210E-3:
                Vrange = "AUTO ON"
            else:
                Vrange = "{:.2f}".format(voltage)
        elif '2450' in self.inst_type:
            if voltage>200:
                Vrange = "AUTO ON"
            elif voltage<210E-3:
                Vrange = "AUTO ON"
            else:
                Vrange = "{:.2f}".format(voltage)
        else:
            Vrange = "AUTO ON"

        #choose Irange
        if '2400' in self.inst_type:
            if current<9E-9:
                Irange = "100E-9"
            elif current<90E-9:
                Irange = "1E-6"
            elif current<9E-6:
                Irange = "10E-6"
            elif current<90E-6:
                Irange = "100E-6"
            elif current<900E-6:
                Irange = "1E-3"
            elif current<9E-3:
                Irange = "10E-3"
            elif current<90E-3:
                Irange = "100E-3"
            elif current<900E-3:
                Irange = "1"
            else:
                Irange = "AUTO ON"
        elif '2440' in self.inst_type:
            if current<9E-9:
                Irange = "100E-9"
            elif current<90E-9:
                Irange = "1E-6"
            elif current<9E-6:
                Irange = "10E-6"
            elif current<90E-6:
                Irange = "100E-6"
            elif current<900E-6:
                Irange = "1E-3"
            elif current<9E-3:
                Irange = "10E-3"
            elif current<90E-3:
                Irange = "100E-3"
            elif current<900E-3:
                Irange = "1"
            else:
                Irange = "AUTO ON"
        elif '2450' in self.inst_type:
            if current<9E-9:
                Irange = "100E-9"
            elif current<90E-9:
                Irange = "1E-6"
            elif current<9E-6:
                Irange = "10E-6"
            elif current<90E-6:
                Irange = "100E-6"
            elif current<900E-6:
                Irange = "1E-3"
            elif current<9E-3:
                Irange = "10E-3"
            elif current<90E-3:
                Irange = "100E-3"
            elif current<900E-3:
                Irange = "1"
            else:
                Irange = "AUTO ON"
        else:
            Irange = "AUTO ON"

        #self.inst.write(f'VOLT{self.chConvert[self.ch]} {abs(voltage):3g}')
        self.inst.write(f':SENS:VOLT:RANG:AUTO ON')
        self.inst.write(f':SOUR:FUNC CURR')
        self.inst.write(f':SOUR:CURR:MODE FIXED')
        if 'AUTO ON' in Irange:
            self.inst.write(f':SOUR:CURR:RANG:AUTO ON')
            #print(":SOUR:CURR:RANG:AUTO ON")
        else:
            self.inst.write(f':SOUR:CURR:RANG {Irange}')
            #print(f':SOUR:CURR:RANG {Irange}')
        self.inst.write(f':SOUR:CURR:LEV {current}')
        self.inst.write(f':SENS:FUNC "VOLT"')
        self.inst.write(f':SENS:VOLT:PROT {voltage}')
        self.inst.write(f':SENS:VOLT:RANG:AUTO ON')

        self._configured = True

        #if 'AUTO ON' in Vrange:
        #    self.inst.write(f':SENS:VOLT:RANG:AUTO ON')
        #else:
        #    self.inst.write(f':SENS:VOLT:RANG {Vrange}')

    # @GInstSetMethod(unit = 'V')
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

        self.inst.write(f':SOUR:VOLT:LEV {voltage}')


    # @GInstSetMethod(unit = 'A')
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

        self.inst.write(f':SOUR:CURR:LEV {current}')



    @GInstOnMethod()
    def outputON(self) :
        """
        power.outputON() -> None
        ================================================================
        [power(channel) output ON]
        :param None:
        :return: None.
        """
        #self.inst.write(f'OUT{self.chConvert[self.ch]} 1')
        self.inst.write(f':OUTP:STAT ON')


    @GInstOffMethod()
    def outputOFF(self) :
        """
        power.outputOFF() -> None
        ================================================================
        [power(channel) output OFF]
        :param None:
        :return: None.
        """
        #self.inst.write(f'OUT{self.chConvert[self.ch]} 0')
        self.inst.write(f':OUTP:STAT OFF')


    @GInstGetMethod(unit = 'V')
    def measureVoltage(self, range=1000, times = 1, res=1000 ):
        """
        power.measureVoltage() -> Voltage
        ================================================================
        [power(channel) measure Voltage]
        :param None:
        :return: Voltage.
        """
        "VOLT:DC"
        #cmd_str = "MEAS:VOLT:DC?"

        valstr = self.inst.query(":OUTPUT?")
        #print("imeas: " + valstr)
        if(valstr=="0"):
            #print("OFF: " + valstr)
            self.inst.write(f':OUTP:STAT ON')

        self.inst.write(f':SENS:FUNC "VOLT"')
        self.inst.write(f':SENS:VOLT:RANG:AUTO ON')
        #self.inst.write(f':FORM:ELEM VOLT')

        cmd_str = ":form:elem volt"
        cnt = 0
        average = 0
        meas_val = 0
        while cnt<times:
            valstr = self.inst.query(cmd_str)
            #print("vmeas: " + valstr)
            meas_val = float(re.search(r"[-+]?\d*\.\d*[Ee][+-]\d+|\d+", valstr).group(0))
            average = average + meas_val/times
            #print("meas_val: " + str(meas_val))
            #print("average: " + str(average))
            cnt = cnt+1
        return average



    @GInstGetMethod(unit = 'A')
    def measureCurrent(self, range=10, times = 1, res=10 ):
        """
        power.measureCurrent() -> Current
        ================================================================
        [power(channel) measure Current]
        :param None:
        :return: Current.
        """
        #cmd_str = "MEAS:CURR:DC?"
        #self.inst.write(f':SENS:CURR:PROT 10E-3')
        #self.inst.write(f':SENS:FUNC "CURR"')
        #self.inst.write(f':SENS:CURR:RANG 10E-3')
        #self.inst.write(f':SENS:CURR:PROT 10E-3')

        valstr = self.inst.query(":OUTPUT?")
        #print("imeas: " + valstr)
        if(valstr=="0"):
            #print("OFF: " + valstr)
            self.inst.write(f':OUTP:STAT ON')

        self.inst.write(f':SENS:FUNC "CURR"')
        self.inst.write(f':SENS:CURR:RANG:AUTO ON')

        cmd_str = ":form:elem curr"

        cnt = 0
        average = 0
        meas_val = 0
        while cnt<times:
            valstr = self.inst.query(cmd_str)
            #print("imeas: " + valstr)
            meas_val = float(re.search(r"[-+]?\d*\.\d*[Ee][+-]\d+|\d+", valstr).group(0))
            average = average + meas_val/times
            #print("meas_val: " + str(meas_val))
            #print("average: " + str(average))
            cnt = cnt+1
        return average


