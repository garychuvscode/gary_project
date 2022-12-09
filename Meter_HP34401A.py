'''
since this library is from Geroge, need to modify before mapped to the testing system
1. need to have the fake GInst in the related folder
2. need to cancel all the deccorator
3. check how to build up the interface to connect two library
'''
import pyvisa
import re

from .GInst import *


class Meter_HP34401A(GInst):


    def __init__(self, link, ch = 0) :
        super().__init__()

        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch

        try :
            self.inst = rm.open_resource(link)
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            self.inst.baud_rate = 38400
        except Exception as e:
            raise Exception(f'<>< Meter_HP34401A ><> open FuncGen Fail {str(e)}!')

        idn = self.inst.query('*IDN?')

        if '34401A' not in idn and '34461A' not in idn and '34465A' not in idn:
            raise Exception(f'<>< Meter_HP34401A ><> 34401A/34461A/34465A not fit in "{idn}" !')


    @GInstGetMethod(unit = 'V')
    def measureVoltage(self, range='DEF', res='DEF', times = 1):
        """
        meter.measureVoltage() -> Voltage
        ================================================================
        [meter(channel) measure Voltage]
        :param None:
        :return: Voltage.
        """
        #print("meas_range: " + str(range))
        #print("meas_res: " + str(res))
        #print("meas_times: " + str(times))

        #print("meas_range_f: " + str(range))
        #print("meas_res_f: " + str(res))
        #print("meas_times_f: " + str(times))
        cmd_str = f"MEAS:VOLT:DC? {range}, {res}"

        cnt = 0
        average = 0
        meas_val = 0
        while cnt<times:
            valstr = self.inst.query(cmd_str)
            #print("vmeas: " + valstr)
            volt = float(re.search(r"[-+]?\d*\.\d*[Ee][+-]\d+|\d+", valstr).group(0))
            average = average + volt/times
            #print("volt: " + str(volt))
            #print("average: " + str(average))
            cnt = cnt+1
        return average


    @GInstGetMethod(unit = 'A')
    def measureCurrent(self, range='DEF', res='DEF', times = 1):
        """
        power.measureCurrent() -> Current
        ================================================================
        [meter(channel) measure Current]
        :param None:
        :return: Current.
        """
        cmd_str = f"MEAS:CURR:DC? {range}, {res}"

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


    @GInstGetMethod(unit = 'ohm')
    def measureResistor(self, range='DEF', res='DEF', times = 1):
        """
        meter.measureResistor() -> Resistor
        ================================================================
        [meter(channel) measure Resistor]
        :param None:
        :return: Resistor.
        """
        cmd_str = f"MEAS:RES? {range}, {res}"

        cnt = 0
        average = 0
        meas_val = 0
        while cnt<times:
            valstr = self.inst.query(cmd_str)
            #print("vmeas: " + valstr)
            volt = float(re.search(r"[-+]?\d*\.\d*[Ee][+-]\d+|\d+", valstr).group(0))
            average = average + volt/times
            #print("volt: " + str(volt))
            #print("average: " + str(average))
            cnt = cnt+1
        return average

