import pyvisa
import re

from .GInst import *


class Loader_Chroma63000(GInst):

    def __init__(self, link, ch = 0) :
        super().__init__()
        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch
        self.chConvert = {'CH1':'1', 'CH2':'2', 'CH3':'3', 'CH4':'4', 'ALL':':1'}

        try :
            self.inst = rm.open_resource(link)
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            self.inst.baud_rate = 38400
        except Exception as e:
            raise Exception(f'<>< Loader_Chroma63000 ><> open Loader Fail {str(e)}!')

        self.inst.write(f'*RST')
        idn = self.inst.query('*IDN?')

        #source meter model#
        if 'CHROMA,63' in idn:
            self.inst_type = "CHROMA,63"
        else:
            raise Exception(f'<>< Loader_Chroma63000 ><> CHROMA,63 not fit in "{idn}"!')
       


    def close(self) :
        print("close Chroma,63xx")

    @GInstSetMethod(unit = 'V')  #decorate used to update UI
    def configureVoltageSource(self, voltage, current) :
        """
        power.setVoltage(voltage) -> None
        ================================================================
        [power(channel) set Voltage]
        :param voltage:
        :return: None.
        """

        cmd_str = f':CHAN:LOAD {self.chConvert[self.ch]}'
        print(cmd_str)
        self.inst.write(cmd_str)

        if self.chConvert[self.ch] == '1':
            cmd_str = f':SHOW:DISP L'
            print(cmd_str)
            self.inst.write(cmd_str)
        elif self.chConvert[self.ch] == '3':
            cmd_str = f':SHOW:DISP L'
            print(cmd_str)
            self.inst.write(cmd_str)
        else:       ##self.ch == 2 or 3
            cmd_str = f':SHOW:DISP R'
            print(cmd_str)
            self.inst.write(cmd_str)

        cmd_str = f':MODE CVH'
        print(cmd_str)
        self.inst.write(cmd_str)

        cmd_str = f':CURRENT:STATIC:L1 {current}'
        #self.inst.write(f':CURRENT:STATIC:L1 {current}')
        print(cmd_str)
        self.inst.write(cmd_str)


    @GInstSetMethod(unit = 'A')
    def configureCurrentSource(self, voltage, current) :
        """
        power.setCurrent(current) -> None
        ================================================================
        [power(channel) set Current]
        :param current:
        :return: None.
        """
        self.inst.write(f':CHAN:LOAD {self.chConvert[self.ch]}')

        if self.chConvert[self.ch] == '1':
            cmd_str = f':SHOW:DISP L'
            print(cmd_str)
            self.inst.write(cmd_str)
        elif self.chConvert[self.ch] == '3':
            cmd_str = f':SHOW:DISP L'
            print(cmd_str)
            self.inst.write(cmd_str)
        else:       ##self.ch == 2 or 3
            cmd_str = f':SHOW:DISP R'
            print(cmd_str)
            self.inst.write(cmd_str)

        self.inst.write(f':MODE CCH')

        self.inst.write(f':CURRENT:STATIC:L1 {current}')

    @GInstSetMethod(unit = 'A')
    def setCurrent(self, current) :
        """
        power.setCurrent(current) -> None
        ================================================================
        [power(channel) set Current]
        :param current:
        :return: None.
        """
        self.inst.write(f':CHAN:LOAD {self.chConvert[self.ch]}')

        if(current<0.2):
            self.inst.write(f':MODE CCL')
        elif(current<2):
            self.inst.write(f':MODE CCM')
        else:
            self.inst.write(f':MODE CCH')

        self.inst.write(f':CURRENT:STATIC:L1 {current}')

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
        self.inst.write(f':CHAN:LOAD {self.chConvert[self.ch]}')
        self.inst.write(f':LOAD ON')

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
        self.inst.write(f':CHAN:LOAD {self.chConvert[self.ch]}')
        self.inst.write(f':LOAD OFF')

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
        #print("meas_range: " + str(range))
        #print("meas_res: " + str(res))
        #print("meas_times: " + str(times))

        self.inst.write(f':CHAN:LOAD {self.chConvert[self.ch]}')

        if self.chConvert[self.ch] == '1':
            cmd_str = f':SHOW:DISP L'
            print(cmd_str)
            self.inst.write(cmd_str)
        elif self.chConvert[self.ch] == '3':
            cmd_str = f':SHOW:DISP L'
            print(cmd_str)
            self.inst.write(cmd_str)
        else:       ##self.ch == 2 or 3
            cmd_str = f':SHOW:DISP R'
            print(cmd_str)
            self.inst.write(cmd_str)

        cmd_str = "MEAS:VOLT?"

        cnt = 0
        average = 0
        volt = 0
        while cnt<times:
            valstr = self.inst.query(cmd_str)
            print("vmeas: " + valstr)
            #volt = float(re.search(r"[-+]?\d*\.\d*[Ee][+-]\d+|\d+", valstr).group(0))
            volt = float(valstr)
            average = average + volt/times
            #print("volt: " + str(volt))
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

        cmd_str = "MEAS:CURR?"

        cnt = 0
        average = 0
        meas_val = 0
        while cnt<times:
            valstr = self.inst.query(cmd_str)
            print("imeas: " + valstr)
            #meas_val = float(re.search(r"[-+]?\d*\.\d*[Ee][+-]\d+|\d+", valstr).group(0))
            meas_val = float(valstr)
            average = average + meas_val/times
            #print("meas_val: " + str(meas_val))
            #print("average: " + str(average))
            cnt = cnt+1
        return average
