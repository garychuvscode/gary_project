
from .GInst import *

RaisedE = False

# 220923 this file is the virtual machine for simulation mode
# import Uni_virtual as Uni
# the way to define: pwr1 = Uni.Uni_Virtual(link, ch)


class Uni_Virtual(GInst):
    VI = 0
    IO = 0
    R = 2  # ohm
    RATIO = 2

    def __init__(self, link, ch):
        super().__init__()

        self.link = link
        self.ch = ch

    # magic function for virtual machine
    def __getattr__(self, name):
        self.name = name
        return self.foo

    @GInstOnMethod()
    def outputON(self):
        pass

    @GInstOffMethod()
    def outputOFF(self):
        pass

    @GInstOnMethod()
    @GInstSetMethod(unit='V')
    def setVoltage(self, V):
        '''
        if V == 4 :
            print("setV:", V)
            global RaisedE
            if RaisedE is False :
                RaisedE = True
                raise IOError
        '''
        Uni_Virtual.VI = V

    @GInstSetMethod(unit='A')
    def setCurrent(self, I):
        Uni_Virtual.IO = I

    @GInstGetMethod(unit='V')
    def measureVoltage(self, range=0):
        if self.ch == 'IN':
            ii = (Uni_Virtual.IO*Uni_Virtual.RATIO)
            return Uni_Virtual.VI - ii * Uni_Virtual.R

        if self.ch == 'OUT':
            return Uni_Virtual.VI * Uni_Virtual.RATIO

        return 0

    @GInstGetMethod(unit='A')
    def measureCurrent(self, range=0):
        if self.ch == 'IN':
            return (Uni_Virtual.IO*Uni_Virtual.RATIO)

        if self.ch == 'OUT':
            return Uni_Virtual.IO

        return 0

    def foo(self, boo=0, **kwarg):

        return boo

    def printScreenToPC(self, path):
        """
        scope.printScreenToPC(path) -> None
        ================================================================
        [print Lecroy Screen and save to path]
        :param path: file path
        :return: None
        """
        from pathlib import PurePath
        from shutil import copyfile
        file = str(PurePath('./') / 'dll' / 'PIC_TestImage.png')

        copyfile(file, path)

    def waitTriggered(self, timeout_ms=5000):
        """
        scope.waitTriggered(timeout = 5) -> bool
        ================================================================
        [wait until scope is triggered]
        :param timeout(option): timeout second
        :return: None
        """
        import time
        time.sleep(0.5)
        return True

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

        print(f"{chV_str=}, {chI_str=}")
