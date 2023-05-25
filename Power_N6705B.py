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
class Power_N6705B(GInst):

    def __init__(self, link, ch):
        super().__init__()

        rm = pyvisa.ResourceManager()

        self.link = link
        self.ch = ch

        self.chConvert = {'CH1': '@1', 'CH2': '@2',
                          'CH3': '@3', 'CH4': '@4', 'ALL': '@1,2,3,4'}

        try:
            self.inst = rm.open_resource(link)
            self.inst.read_termination = '\n'
            self.inst.write_termination = '\n'
            self.inst.baud_rate = 38400
        except Exception as e:
            raise Exception(
                f'<>< Power_N6705B ><> open 4-CH Power Fail {str(e)}!')

        idn = self.inst.query('*IDN?')

        if 'N6705' not in idn:
            raise Exception(
                f'<>< Power_N6705B ><> "N6705" not fit in "{idn}"!')

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

        self.inst.write(f'Volt {abs(voltage):3g},({self.chConvert[self.ch]})')

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

        self.inst.write(f'Curr {abs(current):3g},({self.chConvert[self.ch]})')

    @GInstOnMethod()
    def outputON(self):
        """
        power.outputON() -> None
        ================================================================
        [power(channel) output ON]
        :param None:
        :return: None.
        """
        self.inst.write(f'VOLT:MODE FIX,({self.chConvert[self.ch]})')
        #print(f'VOLT:MODE FIX,({self.chConvert[self.ch]})')
        self.inst.write(f'OUTP ON,({self.chConvert[self.ch]})')

    @GInstOffMethod()
    def outputOFF(self):
        """
        power.outputOFF() -> None
        ================================================================
        [power(channel) output OFF]
        :param None:
        :return: None.
        """
        self.inst.write(f'OUTP OFF,({self.chConvert[self.ch]})')

    @GInstGetMethod(unit='V')
    def measureVoltage(self, *args, **kwargs):
        """
        power.measureVoltage() -> Voltage
        ================================================================
        [power(channel) measure Voltage]
        :param None:
        :return: Voltage.
        """
        valstr = self.inst.query(f'MEAS:VOLT? ({self.chConvert[self.ch]})')
        return float(valstr)

    @GInstGetMethod(unit='A')
    def measureCurrent(self, *args, **kwargs):
        """
        power.measureCurrent() -> Current
        ================================================================
        [power(channel) measure Current]
        :param None:
        :return: Current.
        """
        valstr = self.inst.query(f'MEAS:CURR? ({self.chConvert[self.ch]})')
        return float(valstr)

    @GInstOffMethod()
    def abortARB(self):
        """
        power.abortARB(None) -> None
        ================================================================
        [power(channel) abort ARB ]
        :param None
        :return: None.
        """
        self.inst.write(f'ABOR:TRAN (@1,2,3,4)')

    @GInstOnMethod()
    def triggerARB(self, ch1=None, ch2=None, ch3=None, ch4=None):
        """
        power.triggerARB(ch1 = None, ch2 = None, ch3 = None, ch4 = None) -> Current
        ================================================================
        [power(channel) trigger ARB ]
        :param ch1, ch2, ch3, ch4: given True for Voltage else given string 'I' for Current to enable channel ARB
        :return: None.
        """
        ch1 = True if ch1 is None and self.ch == 'CH1' else ch1
        ch2 = True if ch2 is None and self.ch == 'CH2' else ch2
        ch3 = True if ch3 is None and self.ch == 'CH3' else ch3
        ch4 = True if ch4 is None and self.ch == 'CH4' else ch4

        ch_str = f"@{'1,' if ch1 else ''}{'2,' if ch2 else ''}{'3,' if ch3 else ''}{'4,' if ch4 else ''}"[
            :-1]
        chV_str = f"@{'1,' if ch1 else ''}{'2,' if ch2 else ''}{'3,' if ch3 else ''}{'4,' if ch4 else ''}"[
            :-1]
        chI_str = f"@{'1,' if isinstance(ch1, str) else ''}{'2,' if isinstance(ch2, str) else ''}{'3,' if isinstance(ch3, str) else ''}{'4,' if isinstance(ch4, str) else ''}"[
            :-1]

        #self.inst.write(f'ARB:FUNC:TYPE VOLT,({ch_str})')
        if len(chV_str) > 0:
            self.inst.write(f'ARB:FUNC:TYPE VOLT,({chV_str})')
            self.inst.write(f'VOLT:MODE ARB,({chV_str})')
        if len(chI_str) > 0:
            self.inst.write(f'ARB:FUNC:TYPE CURR,({chI_str})')
            self.inst.write(f'CURR:MODE ARB,({chI_str})')

        self.inst.write(f'TRIG:ARB:SOUR BUS')
        self.inst.write(f'OUTP ON,({ch_str})')
        self.inst.write(f'INIT:TRAN ({ch_str})')
        self.inst.write(f'*TRG')

    # @GInstSetMethod(unit = 'V', value2=True)

    def setTrapezoid(self, V0=0, V1=0, t0=0.1, t1=1e-3, t2=0.1, t3=1e-3, t4=0.1):
        """
        power.setTrapezoid(V0 = 0, V1 = 0, t0 = 0.1, t1 = 1e-3, t2 = 0.1, t3 = 1e-3, t4 = 0.1) -> None
        ================================================================
        [power(channel) set ARB Trapezoid]
        :param V0,V1,t0,t1,t2,t3,t4: 4-CH power N6705 ARB trap
        :return: None.
        """
        self.inst.write(f"VOLT:SLEW:MAX 1,({self.chConvert[self.ch]})")

        self.inst.write(f'ARB:FUNC:SHAP TRAP,({self.chConvert[self.ch]})')

        self.inst.write(
            f'ARB:VOLT:TRAP:STAR {V0:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:TRAP:TOP  {V1:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:TRAP:STAR:TIM {t0:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:TRAP:RTIM {t1:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:TRAP:TOP:TIM {t2:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:TRAP:FTIM {t3:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:TRAP:END:TIM {t4:.6f},({self.chConvert[self.ch]})')

    def setTrapezoidI(self, I0=0, I1=0, t0=0.1, t1=1e-3, t2=0.1, t3=1e-3, t4=0.1):
        """
        power.setTrapezoidI(I0 = 0, I1 = 0, t0 = 0.1, t1 = 1e-3, t2 = 0.1, t3 = 1e-3, t4 = 0.1) -> None
        ================================================================
        [power(channel) set ARB Trapezoid]
        :param I0,I1,t0,t1,t2,t3,t4: 4-CH power N6705 ARB trap
        :return: None.
        """
        self.inst.write(f"CURR:SLEW:MAX 1,({self.chConvert[self.ch]})")

        self.inst.write(f'ARB:FUNC:SHAP TRAP,({self.chConvert[self.ch]})')

        self.inst.write(
            f'ARB:CURR:TRAP:STAR {I0:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:CURR:TRAP:TOP  {I1:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:CURR:TRAP:STAR:TIM {t0:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:CURR:TRAP:RTIM {t1:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:CURR:TRAP:TOP:TIM {t2:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:CURR:TRAP:FTIM {t3:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:CURR:TRAP:END:TIM {t4:.6f},({self.chConvert[self.ch]})')

    # @GInstSetMethod(unit = 'V', value2=True)

    def setRamp(self, V0=0, V1=0, t0=0.1, t1=1e-3, t2=0.1):
        """
        power.setRamp(V0 = 0, V1 = 0, t0 = 0.1, t1 = 1e-3, t2 = 0.1) -> None
        ================================================================
        [power(channel) set ARB Ramp]
        :param V0,V1,t0,t1,t2: 4-CH power N6705 ARB Ramp
        :return: None.
        """
        self.inst.write(
            f'ARB:TERM:LAST ON,({self.chConvert[self.ch]})')    # keep last voltage
        self.inst.write(f'ARB:FUNC:SHAP RAMP,({self.chConvert[self.ch]})')

        self.inst.write(
            f'ARB:VOLT:RAMP:STAR {V0:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:RAMP:END  {V1:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:RAMP:STAR:TIM {t0:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:RAMP:RTIM {t1:.6f},({self.chConvert[self.ch]})')
        self.inst.write(
            f'ARB:VOLT:RAMP:END:TIM {t2:.6f},({self.chConvert[self.ch]})')


if __name__ == '__main__':

    ch = 'CH1'

    ch1 = True
    ch2 = False
    ch3 = True
    ch4 = 'I'

    ch1 = True if ch1 is None and ch == 'CH1' else ch1
    ch2 = True if ch2 is None and ch == 'CH2' else ch2
    ch3 = True if ch3 is None and ch == 'CH3' else ch3
    ch4 = True if ch4 is None and ch == 'CH4' else ch4

    ch_str = f"@{'1,' if ch1 else ''}{'2,' if ch2 else ''}{'3,' if ch3 else ''}{'4,' if ch4 else ''}"[
        :-1]
    chV_str = f"@{'1,' if ch1 is True else ''}{'2,' if ch2 is True else ''}{'3,' if ch3 is True else ''}{'4,' if ch4 is True else ''}"[
        :-1]
    chI_str = f"@{'1,' if isinstance(ch1, str) else ''}{'2,' if isinstance(ch2, str) else ''}{'3,' if isinstance(ch3, str) else ''}{'4,' if isinstance(ch4, str) else ''}"[
        :-1]

    print(f"ch_str  = {ch_str}")
    print(f"chV_str = {chV_str}")
    print(f"chI_str = {chI_str}")
