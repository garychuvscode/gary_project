import pyvisa
import re

from .GInst import *

# turn off the formatter
# fmt: off
class FuncGen_AFG31000(GInst):

    def __init__(self, link, ch = 0) :
        super().__init__()
        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch
        self.chConvert = {'CH1':'1', 'CH2':'2', 'ALL':':1'}

        self.PwmByPulseMode = False
        self.PwmByPulseFreq = 1000

        try :
            self.inst = rm.open_resource(link)
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            self.inst.baud_rate = 38400
        except Exception as e:
            raise Exception(f'<>< FuncGen_AFG31000 ><> open FuncGen Fail {str(e)}!')

        self.inst.write(f'*RST')
        idn = self.inst.query('*IDN?')

        #source meter model#
        if 'AFG3' in idn:
            self.inst_type = "AFG3x"
        else:
            raise Exception(f'<>< FuncGen_AFG31000 ><> AFG31 not fit in "{idn}"!')


    def outpIMP(self, set_val) :
        """
        power.setCurrent(current) -> None
        ================================================================
        [power(channel) set Current]
        :param current:
        :return: None.
        """

        #print(f'set_val {set_val}')
        if(set_val == '50'):
            cmd_str = f'OUTP{self.chConvert[self.ch]}:IMP {set_val}'
        elif(set_val == 'INF'):
            cmd_str = f'OUTP{self.chConvert[self.ch]}:IMP {set_val}'
        elif(set_val == 'MIN'):
            cmd_str = f'OUTP{self.chConvert[self.ch]}:IMP {set_val}'
        elif(set_val == 'MAX'):
            cmd_str = f'OUTP{self.chConvert[self.ch]}:IMP {set_val}'
        else:
            cmd_str = f'OUTP{self.chConvert[self.ch]}:IMP {50}'
        print(cmd_str)
        self.inst.write(cmd_str)


    def setFuncShape(self, set_val) :
        """
        funcgen.setFuncShape(set_val) -> None
        ================================================================
        [funcgen(channel) set FuncShape]
        :param set_val: "SIN", "SQU",  "RAMP",  "PULS",  "DC",
        :return: None.
        """
        if(set_val == "SIN"):
            cmd_str = f':SOUR{self.chConvert[self.ch]}:FUNC:SHAPE {set_val}'
        elif(set_val == "SQU"):
            cmd_str = f':SOUR{self.chConvert[self.ch]}:FUNC:SHAPE {set_val}'
        elif(set_val == "RAMP"):
            cmd_str = f':SOUR{self.chConvert[self.ch]}:FUNC:SHAPE {set_val}'
        elif(set_val == "PULSE"):
            cmd_str = f':SOUR{self.chConvert[self.ch]}:FUNC:SHAPE {set_val}'
        else:
            cmd_str = f':SOUR{self.chConvert[self.ch]}:FUNC:SHAPE DC'

        self.inst.write(cmd_str)

    # def setFreq(self, set_val) :
    #     """
    #     power.setVoltage(voltage) -> None
    #     ================================================================
    #     [power(channel) set Voltage]
    #     :param voltage:
    #     :return: None.
    #     """

    #     cmd_str = f':SOUR{self.chConvert[self.ch]}:FREQ {set_val}'
    #     print(cmd_str)
    #     self.inst.write(cmd_str)

    @GInstGetMethod(unit = 'Hz')
    def setFreq(self, frquence):
        """
        funcgen.setFreq(frquence) -> None
        ================================================================
        [funcgen(channel) set Frequence]
        :param frquence:
        :return: None.
        """
        if not self.PwmByPulseMode:
            self.inst.write(f"SOUR{self.chConvert[self.ch]}:FUNC:SHAPE PULS\n")
            self.inst.write(f"SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:HIGH 3.3V\n")
            self.inst.write(f"SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:LOW  0.0V\n")

        self.inst.write(f"SOUR{self.chConvert[self.ch]}:FREQ {frquence}\n")
        self.PwmByPulseFreq = frquence

    @GInstGetMethod(unit = '%')
    def setDuty(self, duty):
        """
        funcgen.setDuty(duty) -> None
        ================================================================
        [funcgen(channel) set Duty]
        :param duty: 0-100 for 0-100%
        :return: None.
        """
        if duty < 0.001 :
            self.inst.write(f"SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:HIGH 0.2V\n")
            self.PwmByPulseMode = False

        elif duty > 99.999 :
            self.inst.write(f"SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:LOW 3.1V\n")
            self.PwmByPulseMode = False

        else :
            if not self.PwmByPulseMode:
                self.inst.write(f"SOUR{self.chConvert[self.ch]}:FUNC:SHAPE PULS\n")
                self.inst.write(f"SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:HIGH 3.3V\n")
                self.inst.write(f"SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:LOW  0.0V\n")
                self.PwmByPulseMode = True

            pulse_width = duty / (self.PwmByPulseFreq * 100)
            self.inst.write(f"SOUR{self.chConvert[self.ch]}:PULSE:WIDTH {pulse_width:E}\n")


    def pulsePeriod(self, set_val) :
        """
        funcgen.pulsePeriod(set_val) -> None
        ================================================================
        [funcgen(channel) set pulsePeriod]
        :param set_val:
        :return: None.
        """
        cmd_str = f'SOUR{self.chConvert[self.ch]}:PULS:PER {set_val}'
        self.inst.write(cmd_str)


    def pulseWidth(self, set_val) :
        """
        funcgen.pulseWidth(current) -> None
        ================================================================
        [funcgen(channel) set pulseWidth]
        :param set_val:
        :return: None.
        """
        cmd_str = f'SOUR{self.chConvert[self.ch]}:PULS:WIDT {set_val}'
        self.inst.write(cmd_str)


    def pulseLeading(self, set_val) :
        """
        funcgen.pulseLeading(set_val) -> None
        ================================================================
        [funcgen(channel) set pulseLeading]
        :param set_val:
        :return: None.
        """
        cmd_str = f'SOUR{self.chConvert[self.ch]}:PULS:TRAN:LEAD {set_val}'
        self.inst.write(cmd_str)


    def pulseTrailing(self, set_val) :
        """
        funcgen.pulseTrailing(set_val) -> None
        ================================================================
        [funcgen(channel) set pulseTrailing]
        :param set_val:
        :return: None.
        """
        cmd_str = f'SOUR{self.chConvert[self.ch]}:PULS:TRAN:TRA {set_val}s'
        self.inst.write(cmd_str)


    def voltHigh(self, voltage) :
        """
        funcgen.voltHigh(voltage) -> None
        ================================================================
        [funcgen(channel) set voltage]
        :param current:
        :return: None.
        """
        cmd_str = f'SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:HIGH {voltage}'
        self.inst.write(cmd_str)

    def voltLow(self, voltage) :
        """
        funcgen.voltLow(current) -> None
        ================================================================
        [funcgen(channel) set voltage]
        :param current:
        :return: None.
        """
        cmd_str = f'SOUR{self.chConvert[self.ch]}:VOLT:LEV:IMM:LOW {voltage}'
        self.inst.write(cmd_str)


    @GInstOnMethod()
    def outputON(self) :
        """
        funcgen.outputOFF(current) -> None
        ================================================================
        [funcgen(channel) output ON]
        :param current:
        :return: None.
        """
        cmd_str = f'OUTP{self.chConvert[self.ch]}:STAT 1'
        self.inst.write(cmd_str)

    @GInstOffMethod()
    def outputOFF(self) :
        """
        funcgen.outputOFF() -> None
        ================================================================
        [funcgen(channel) output OFF]
        :param None:
        :return: None.
        """
        cmd_str = f'OUTP{self.chConvert[self.ch]}:STAT 0'
        self.inst.write(cmd_str)
