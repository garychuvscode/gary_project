
import win32com.client  # import the pywin32 library
import pythoncom
import logging
import os

from GInst import *
# from .GInst import *
'''
the reason not to use .GInst is because all the file now is at the root of python
and it should be different when source code is in different folder
.GInst means GInst.py is at the same folder with current python file
it's a better way to import without folder problems
'''


class Scope_LE6100A(GInst):
    ActiveDSO = None
    '''
    link is to connect for active DSO: 'GPIB: (address)' \n
    or using IP address: 'IP: (address)' \n
    ch is no used in scope application
    '''

    def __init__(self, link, ch, sim_inst0, excel0):
        # this is the function for GInst, call the initial of GInst
        # super is the function using dad's function
        # call the initial of GInst
        super().__init__()
        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst
            import parameter_load_obj as par
            # for the jump out window

            # initial the object and set to simulation mode

            # using the main control book as default
            excel0 = par.excel_parameter('obj_main')
            # ======== only for object programming

        self.link = link
        self.ch = ch
        self.excel_s = excel0

        # extra code added to the instrument object
        # simulation mode setting of the insturment
        self.sim_inst = sim_inst0

        # this is the index number for the temp saving piture
        self.pic_index = 0

        if self.sim_inst == 1:

            logging.debug(f'Initialize LecroyActiveDSO link={link}, ch={ch}')

            if Scope_LE6100A.ActiveDSO is None:
                pythoncom.CoInitialize()
                # .com component connection standard code (for win32com code)
                Scope_LE6100A.ActiveDSO = win32com.client.Dispatch(
                    "LeCroy.ActiveDSOCtrl.1")
                pythoncom.CoInitialize()

            self.inst = Scope_LE6100A.ActiveDSO
            r = self.inst.MakeConnection(link)
            logging.debug(f"self.inst.MakeConnection(link) = {r}")

            if r == 0:
                raise Exception(
                    '<>< Scope_LE6100A ><> LecroyActiveDSO link fail.')

            self.inst.SetRemoteLocal(1)
            self.inst.SetTimeout(20)

            pass
        else:
            print('simulation mode of the scope, initialize function')

            pass

        print('initial of scope object finished')
        pass

    def writeVBS(self, cmd):
        """
        scope.writeVBS(cmd) -> None
        ================================================================
        [write Lecroy VBS-Command to Scope]
        :param cmd: lecory VBS command
        :return: None.
        """

        if self.sim_inst == 1:
            logging.debug(f': {cmd}')
            self.inst.WriteString(f"VBS '{cmd}'", 1)

            if 0 == self.inst.WaitForOPC():
                raise Exception('LecroyActiveDSO writeVBS timeout or fail.')

            pass
        else:
            print('simulation mode of the scope, write VBS')

            pass

    def readVBS(self, cmd):
        """
        scope.readVBS(cmd) -> string
        ================================================================
        [read string after write Lecroy VBS-Command to Scope]
        :param cmd: lecory VBS command
        :return: scope return string.
        """

        if self.sim_inst == 1:
            logging.debug(f': {cmd}')
            self.inst.WriteString(f"VBS? '{cmd}'", 1)

            return self.inst.ReadString(256)  # reads a maximum of 256 bytes

        else:
            print('simulation mode of the scope, read VBS')

            pass

        pass

    def readVBS_float(self, cmd):
        """
        scope.readVBS_float(cmd) -> float
        ================================================================
        [ get float value after write Lecroy VBS-Command to Scope]
        :param cmd: lecory VBS command
        :return: float value.
        """

        if self.sim_inst == 1:
            try:
                return float(self.readVBS(cmd))
            except:
                return -999

        else:
            print('simulation mode of the scope, read VBS, float')

            pass

        pass

    def printScreenToPC(self, path=0):
        # 0 is using the default path

        if path == 0.5:
            # this is the sim_mode for testing
            path = 'c:\\py_gary\\test_excel\\test_pic.png  rf'
            pure_path = os.path.splitext(path)[0]
            # . or 'space' will be cut and neglect
            # only path left
            print(pure_path)
            # since there are already PNG in below command, no need to add '.png'
            pass

        elif path == 0:

            pure_path = self.excel_s.wave_path + self.excel_s.wave_condition

            pass
        """
        scope.printScreenToPC(path) -> None
        ================================================================
        [print Lecroy Screen and save to path]
        :param path: file path
        :return: None
        """
        if self.sim_inst == 1:

            # tell scope to save waveform
            # need another function for loaded the pictrue to the excel
            # input picture to excel, using xlwings
            # sheet.picture.add()
            # path: C://test//picture.png
            # need to be PNG (or will have error)
            pure_path = os.path.splitext(path)[0]
            # full path is used to connect to the excel and load the capture to
            # excel
            full_path = pure_path + '.png'
            self.excel_s.full_path = full_path

            self.inst.StoreHardcopyToFile(
                "PNG", "BCKG,WHITE,AREA,GRIDAREAONLY", full_path)

            pass
        else:

            print('simulation mode of the scope, printScreenToPC')
            pass

        pass

    def triggerSingle(self):

        if self.sim_inst == 1:
            #logging.debug(f': {cmd}')
            #self.inst.WriteString(f"VBS '{cmd}'", 1)
            #self.writeVBS('app.Acquisition.TriggerMode = "Stopped"')
            self.writeVBS('app.Acquisition.TriggerMode = "Single"')
            self.writeVBS('app.ClearSweeps')

            #self.inst.WriteString(f"VBS 'TriggerDetected = app.Acquisition.Acquire({timeout}, false)'", 1)
            # if 0 == self.inst.WaitForOPC():
            #    raise Exception('LecroyActiveDSO writeVBS timeout or fail.')
            pass
        else:
            print('sopce sim mode, trigger single')
            pass

        pass

    def waitTriggered(self, timeout_ms=5000):
        """
        scope.waitTriggered(timeout = 5) -> bool
        ================================================================
        [wait until scope is triggered]
        :param timeout(option): timeout second
        :return: None
        """

        if self.sim_inst == 1:
            import time
            interval_ms = 330

            for t in range(0, timeout_ms, interval_ms):
                if self.readVBS(f'return = app.Acquisition.TriggerMode') == 'Stopped':
                    return True
                time.sleep(interval_ms/1000)

            return False
            # if 0 == self.inst.WaitForOPC():
            #     raise Exception('LecroyActiveDSO acquire fail.')

            # r = int(self.readVBS(f'return = TriggerDetected'))

            # if r :
            #     return True
            # else :
            #     raise Exception("LE6100A waitTriggered timeout!")
            pass

        else:
            print('sim mode for scope, wait for trigger')
            pass

        pass

    def getFitScale(self, value, in_grid=8):
        """
        scope.getFitScale(value, in_grid) -> scale(float)
        ================================================================
        [get fit scale to make value view in in_grid]
        :param value:
        :param in_grid:
        :return: None
        """

        if self.sim_inst == 1:

            value_one_grid = value / in_grid

            # E = M * 10^N
            vestr = f'{value_one_grid:e}'
            ep = vestr.rfind("e")

            M = float(vestr[:ep])
            N = float(vestr[ep+1:])

            print(f"vestr={vestr}, M={M}, N={N}")
            if M <= 1:
                vM = 1
            elif M <= 2:
                vM = 2
            elif M <= 5:
                vM = 5
            else:
                vM = 10

            return vM * (10 ** N)
            pass
        else:
            print('sim mode for scope, find fit scale')

    def scope_initial(self):
        '''
        socpe initialize function, plan to add loading setting file in scope in the future
        '''
        if self.sim_inst == 1:
            # run initial settings

            pass
        else:
            print('sim mode scope, initial settings')
            pass

        pass


if __name__ == '__main__':
    #  the testing code for this file object
    import parameter_load_obj as par
    excel_t = par.excel_parameter('obj_main')
    sim_scope = 0

    scope = Scope_LE6100A('GPIB: 5', 3, sim_scope, excel_t)
    scope.scope_initial()
    # testing for the scope capture

    scope.printScreenToPC(path=0.5)
