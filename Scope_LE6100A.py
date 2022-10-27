

from GInst import *
import win32com.client  # import the pywin32 library
import pythoncom
import logging
import os
import time
import pyvisa


# maybe the instrument need the delay time
rm = pyvisa.ResourceManager()

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
        # if = 1, scope object is already link to the real scope
        self.link_status = 0

        # extra code added to the instrument object
        # simulation mode setting of the insturment
        self.sim_inst = sim_inst0

        # the information for GPIB resource manager
        self.GP_addr_ini = self.excel_s.scope_addr

        # the setting for each channel
        # to be update in the function: ini_setting_select
        self.ch_c1 = {}
        self.ch_c2 = {}
        self.ch_c3 = {}
        self.ch_c4 = {}
        self.ch_c5 = {}
        self.ch_c6 = {}
        self.ch_c7 = {}
        self.ch_c8 = {}

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
            self.link_status = 1

            pass
        else:
            print('simulation mode of the scope, initialize function')

            pass

        print('initial of scope object finished')
        pass

    def open_inst(self):

        if self.link == 0:
            logging.debug(
                f'Initialize LecroyActiveDSO link={self.link}, ch={self.ch}')

            if Scope_LE6100A.ActiveDSO is None:
                pythoncom.CoInitialize()
                # .com component connection standard code (for win32com code)
                Scope_LE6100A.ActiveDSO = win32com.client.Dispatch(
                    "LeCroy.ActiveDSOCtrl.1")
                pythoncom.CoInitialize()

            self.inst = Scope_LE6100A.ActiveDSO
            r = self.inst.MakeConnection(self.link)
            logging.debug(f"self.inst.MakeConnection(link) = {r}")

            if r == 0:
                raise Exception(
                    '<>< Scope_LE6100A ><> LecroyActiveDSO link fail.')

            self.inst.SetRemoteLocal(1)
            self.inst.SetTimeout(20)
            self.link_status = 1

            pass
        else:
            if self.sim_inst == 0:
                print('call scope open inst in simulation mode')

            else:
                print('scope is already open')

            pass

        print('GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        if self.sim_inst == 1:
            self.inst_obj = rm.open_resource(
                'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
            time.sleep(0.05)
            pass
        else:
            print('now is open the power supply, in address: ' +
                  str(int(self.GP_addr_ini)))
            # in simulation mode, inst_obj need to be define for the simuation mode
            self.inst_obj = 'power supply simulation mode object'
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
            # print('simulation mode of the scope, write VBS')
            print('write, command is : ' + str(cmd))

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
            # print('simulation mode of the scope, read VBS')
            print('read, command is : ' + str(cmd))

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
            # print('simulation mode of the scope, read VBS, float')
            print('read_float, command is : ' + str(cmd))

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
            if self.sim_inst == 0:
                pure_path = self.excel_s.wave_path + self.excel_s.wave_condition

            else:
                path = self.excel_s.wave_path + self.excel_s.wave_condition
                # replace all the point to p prevent error of the name
                path = path.replace('.', 'p')

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

    def ch_default_setting(self):

        # setup channel 1 to 8
        temp_dict = {}
        gen_dict = self.set_general

        for i in range(1, 8+1):
            # assign the mapped dictionary
            if i == 1:
                temp_dict = self.ch_c1
            elif i == 2:
                temp_dict = self.ch_c2
            elif i == 3:
                temp_dict = self.ch_c3
            elif i == 4:
                temp_dict = self.ch_c4
            elif i == 5:
                temp_dict = self.ch_c5
            elif i == 6:
                temp_dict = self.ch_c6
            elif i == 7:
                temp_dict = self.ch_c7
            elif i == 8:
                temp_dict = self.ch_c8

            self.writeVBS(
                f'app.Acquisition.C{i}.View = {temp_dict["ch_view"]}')
            self.writeVBS(
                f'app.Acquisition.C{i}.VerScale = {temp_dict["volt_dev"]}')
            self.writeVBS(
                f'app.Acquisition.C{i}.BandwidthLimit = {temp_dict["BW"]}')
            self.writeVBS(
                f'app.Acquisition.C{i}.EnhanceResTypes = "{temp_dict["filter"]}"')
            self.writeVBS(
                f'app.Acquisition.C{i}.Veroffset = {temp_dict["v_offset"]}')
            self.writeVBS(
                f'app.Acquisition.C{i}.LabelsText = "{temp_dict["label_name"]}"')
            self.writeVBS(
                f'app.Acquisition.C{i}.LabelsPosition = "{temp_dict["label_position"]}"')
            self.writeVBS(
                f'app.Acquisition.C{i}.ViewLabels = {temp_dict["label_view"]}')
            self.writeVBS(
                f'app.Acquisition.C{i}.Coupling = "{temp_dict["coupling"]}"')

            if self.sim_inst == 0:
                print(f'simulatiuon mode setting the scope channel{i}')

        # setting of trigger
        self.writeVBS(
            f'app.Acquisition.TriggerMode = "{gen_dict["trigger_mode"]}"')
        self.writeVBS(
            f'app.Acquisition.Trigger.Source = "{gen_dict["trigger_source"]}"')
        self.writeVBS(
            f'app.Acquisition.Trigger.{gen_dict["trigger_source"]}.Level = {gen_dict["trigger_level"]}')
        self.writeVBS(
            f'app.Acquisition.Trigger.{gen_dict["trigger_source"]}.Slope = "{gen_dict["trigger_slope"]}"')

        # setting of x-axis
        self.writeVBS(
            f'app.Acquisition.Horizontal.HorScale = {gen_dict["time_scale"]}')
        self.writeVBS(
            f'app.Acquisition.Horizontal.HorOffset = {gen_dict["time_offset"]}')

        # setting of sample mode
        self.writeVBS(
            f'app.Acquisition.Horizontal.SampleMode = "{gen_dict["sample_mode"]}"')

        if self.sim_inst == 0:
            print('setting up general and trigger setting in scope @ simulation mode')

        pass

    def inst_name(self):
        # get the insturment name
        self.cmd_str_name = ''
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst_obj.query(self.cmd_str_name)
            time.sleep(0.05)

            pass
        else:
            # for the simulatiom mode of change output
            print('check the instrument name, sim mode ')
            print(str(self.cmd_str_name))
            self.in_name = 'scope is in sim mode'

            pass

        return self.in_name

    def scope_initial(self, setup_index):
        '''
        socpe initialize function, plan to add loading setting file in scope in the future
        '''
        # this function used to choose different setup for different index

        # should recall general setup in the scope first to makesure the items not update in the setup selection be correct

        if setup_index == 0:
            # the setting for ripple verification
            self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c3 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c4 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c5 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c6 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c7 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
            self.ch_c8 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': '20MHz', 'filter': '2bits', 'v_offset': 0,
                          'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}

            # setting of general
            self.set_general = {'trigger_mode': 'Auto', 'trigger_source': 'C1', 'trigger_level': '0.5',
                                'trigger_slope': 'Positive', 'time_scale': '0.001', 'time_offset': '-0.003', 'sample_mode': 'RealTime'}

        # call the 'ch_default_setting' for each channel setting and 'scope_initial'
        self.ch_default_setting()

    def trigger_adj(self, mode=None, source=None, level=None, slope=None):
        # this function is used to adjust the trigger function
        '''
        choose which trigger to adjust: \n
        mode: 'Auto', 'Normal', 'Single', 'Stopped' \n
        source: 'Cx' \n
        level: 'number' \n
        slope: 'Positive', 'Negative' \n
        '''

        if mode != None:
            self.writeVBS(
                f'app.Acquisition.TriggerMode = "{mode}"')
            # also change the setting after setting updated
            self.set_general["trigger_mode"] = mode
            pass
        if source != None:
            self.writeVBS(
                f'app.Acquisition.Trigger.Source = "{source}"')
            # also change the setting after setting updated
            self.set_general["trigger_source"] = source
            pass
        if level != None and source == None:
            self.writeVBS(
                f'app.Acquisition.Trigger.{self.set_general["trigger_source"]}.Level = {level}')
            # also change the setting after setting updated
            self.set_general["trigger_source"] = level
            pass
        elif level != None and source != None:
            self.writeVBS(
                f'app.Acquisition.Trigger.{source}.Level = {level}')
            # also change the setting after setting updated
            self.set_general["trigger_level"] = level
            self.set_general["trigger_source"] = source

            pass
        if slope != None and source == None:
            self.writeVBS(
                f'app.Acquisition.Trigger.{self.set_general["trigger_source"]}.Slope = "{slope}"')
            # also change the setting after setting updated
            self.set_general["trigger_source"] = slope
            pass
        elif slope != None and source != None:
            self.writeVBS(
                f'app.Acquisition.Trigger.{source}.Slope = "{slope}"')
            # also change the setting after setting updated
            self.set_general["trigger_slope"] = slope
            self.set_general["trigger_source"] = source

            pass

        pass

    def Hor_scale_adj(self, div=None, offset=None):
        # this program used to adjust the time scale
        'to adjust time scale, enter (/div, offset)'

        if div != None:
            # change the time scale for the scope
            self.writeVBS(
                f'app.Acquisition.Horizontal.HorScale = {div}')
            self.set_general['time_scale'] = div

        if offset != None:
            # hcange the offset of zero point in the scope
            self.writeVBS(
                f'app.Acquisition.Horizontal.HorOffset = {offset}')
            self.set_general['time_offset'] = offset

        pass

    def single_ch_change(self, ch, ver_scale=None, ver_offset=None):

        if ver_scale != None:
            # change the ver div
            self.writeVBS(
                f'app.Acquisition.{ch}.VerScale = {ver_scale}')
            print(f'change {ch} to {ver_scale}')

        if ver_offset != None:
            # change ver offset
            self.writeVBS(
                f'app.Acquisition.{ch}.Veroffset = {ver_offset}')
            print(f'change {ch} to {ver_offset}')

        pass


if __name__ == '__main__':
    #  the testing code for this file object
    import parameter_load_obj as par
    excel_t = par.excel_parameter('obj_main')
    sim_scope = 0
    default_path = 'C:\\py_gary\\test_excel\\wave_form_raw\\'

    scope = Scope_LE6100A('GPIB: 5', 3, sim_scope, excel_t)

    test_index = 2

    if test_index == 0:

        # testing for the scope capture

        scope.printScreenToPC(path=0.5)

        capture = 5
        x = 0
        while x < capture:

            # input path need to include the file name
            final_path = default_path + str(x)
            scope.printScreenToPC(path=final_path)

            x = x + 1
            pass

    elif test_index == 1:
        # teting for the initial setting of scope

        scope.scope_initial(0)
        scope.open_inst()
        scope.inst_name()
        scope.printScreenToPC()
        scope.triggerSingle()
        scope.waitTriggered()
        scope.trigger_adj('Auto', 'C9', '0.56', 'Positive')
        scope.Hor_scale_adj('div_test', 'offset_test')
        scope.single_ch_change('C9', 'ver_div_test', 'ver_off_test')

        pass

    elif test_index == 2:
        # real testing with scope

        scope.scope_initial(0)
        scope.open_inst()
        temp_name = scope.inst_name()
        print(temp_name)
        scope.printScreenToPC(0.5)
        scope.triggerSingle()
        scope.waitTriggered()
        scope.trigger_adj('Stopped', 'C3', '1', 'Negative')
        scope.Hor_scale_adj('0.001', '0.00002')
        scope.single_ch_change('C2', '2', '0.25')
