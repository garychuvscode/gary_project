from GInst import *
import win32com.client  # import the pywin32 library
import pythoncom
import logging
import os
import time
import pyvisa


import scope_set_index as sc_set

# prevent the issue with no control sheet
import sheet_ctrl_main_obj as sh
import parameter_load_obj as para

# fmt: off

# maybe the instrument need the delay time
rm = pyvisa.ResourceManager()

# from .GInst import *
"""
the reason not to use .GInst is because all the file now is at the root of python
and it should be different when source code is in different folder
.GInst means GInst.py is at the same folder with current python file
it's a better way to import without folder problems
"""


class Scope_LE6100A(GInst):
    ActiveDSO = None
    """
    link is to connect for active DSO: 'GPIB: (address)' \n
    or using IP address: 'IP: (address)' \n
    ch is no used in scope application
    """

    def __init__(self, excel0, link="", ch=0, main_off_line0=0):
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
            excel0 = par.excel_parameter("obj_main")
            # ======== only for object programming
        self.excel_s = excel0
        link = "GPIB: " + str(self.excel_s.scope_addr)
        self.link = link
        self.ch = ch

        # if = 1, scope object is already link to the real scope
        self.link_status = 0

        # the information for GPIB resource manager
        self.GP_addr_ini = self.excel_s.scope_addr
        """
        for the GPIB address should be able to loaded from the excel input to the
        instrument, and open automatically
        """

        if self.GP_addr_ini != 100 and main_off_line0 == 0:
            self.sim_inst = 1
        else:
            self.sim_inst = 0

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

        self.nor_v_off = 0
        # the index control variable of the ver_offset

        self.set_general = {
            "trigger_mode": "Auto",
            "trigger_source": "C3",
            "trigger_level": "-3.2",
            "trigger_slope": "Positive",
            "time_scale": "0.0001",
            "time_offset": "-0.0004",
            "sample_mode": "RealTime",
            "fixed_sample_rate": "1.25GS/s",
        }

        self.sc_config = sc_set.scope_config()

        self.g_string_a = [
            "'s smile is attractive.",
            "'s hair is beautiful.",
            "'s face is cute.",
            "'s eyes are shining.",
            " usually comes at 9:30 XD",
            " likes to play stock. ",
            " hates to work overtime. ",
            " is a really good girl XD",
        ]

        # if self.sim_inst == 1:

        #     logging.debug(f'Initialize LecroyActiveDSO link={link}, ch={ch}')

        #     if Scope_LE6100A.ActiveDSO is None:
        #         pythoncom.CoInitialize()
        #         # .com component connection standard code (for win32com code)
        #         Scope_LE6100A.ActiveDSO = win32com.client.Dispatch(
        #             "LeCroy.ActiveDSOCtrl.1")
        #         pythoncom.CoInitialize()

        #     self.inst = Scope_LE6100A.ActiveDSO
        #     r = self.inst.MakeConnection(link)
        #     logging.debug(f"self.inst.MakeConnection(link) = {r}")

        #     if r == 0:
        #         raise Exception(
        #             '<>< Scope_LE6100A ><> LecroyActiveDSO link fail.')

        #     self.inst.SetRemoteLocal(1)
        #     self.inst.SetTimeout(20)
        #     self.link_status = 1

        #     pass
        # else:
        #     print('simulation mode of the scope, initialize function')

        #     pass

        print("initial of scope object finished")
        # self.open_inst()

        # 230503 add the scope capture index for saving picture, default 0
        self.capture_index = 0

        pass

    def open_inst(self):
        if self.sim_inst == 1:
            logging.debug(f"Initialize LecroyActiveDSO link={self.link}, ch={self.ch}")

            if Scope_LE6100A.ActiveDSO is None:
                pythoncom.CoInitialize()
                # .com component connection standard code (for win32com code)
                Scope_LE6100A.ActiveDSO = win32com.client.Dispatch(
                    "LeCroy.ActiveDSOCtrl.1"
                )
                pythoncom.CoInitialize()

            self.inst = Scope_LE6100A.ActiveDSO
            r = self.inst.MakeConnection(self.link)
            logging.debug(f"self.inst.MakeConnection(link) = {r}")

            if r == 0:
                raise Exception("<>< Scope_LE6100A ><> LecroyActiveDSO link fail.")

            self.inst.SetRemoteLocal(1)
            self.inst.SetTimeout(20)
            self.link_status = 1

            # also turn on the resource manager for IDN or other string operation
            self.inst_obj = rm.open_resource(
                "GPIB0::" + str(int(self.GP_addr_ini)) + "::INSTR"
            )

            pass
        else:
            print("call scope open inst")

            pass
        print("GPIB0::" + str(int(self.GP_addr_ini)) + "::INSTR")

        # if self.sim_inst == 1:
        #     self.inst_obj = rm.open_resource(
        #         'GPIB0::' + str(int(self.GP_addr_ini)) + '::INSTR')
        #     time.sleep(0.05)
        #     pass
        # else:
        #     print('now is open the power supply, in address: ' +
        #           str(int(self.GP_addr_ini)))
        #     # in simulation mode, inst_obj need to be define for the simuation mode
        #     self.inst_obj = 'power supply simulation mode object'
        #     pass

    def writeVBS(self, cmd):
        """
        scope.writeVBS(cmd) -> None
        ================================================================
        [write Lecroy VBS-Command to Scope]
        :param cmd: lecory VBS command
        :return: None.
        """

        if self.sim_inst == 1:
            logging.debug(f": {cmd}")
            self.inst.WriteString(f"VBS '{cmd}'", 1)
            # time.sleep(0.03)

            if 0 == self.inst.WaitForOPC():
                raise Exception("LecroyActiveDSO writeVBS timeout or fail.")

            pass
        else:
            # print('simulation mode of the scope, write VBS')
            print("write, command is : " + str(cmd))

            pass

    def readVBS(self, cmd):
        """
        scope.readVBS(cmd) -> string
        ================================================================
        [read string after write Lecroy VBS-Command to Scope]
        :param cmd: lecory VBS command
        :return: scope return string. \n
        example: mode_temp = self.readVBS(f'return = app.Acquisition.TriggerMode')
        """

        if self.sim_inst == 1:
            logging.debug(f": {cmd}")
            self.inst.WriteString(f"VBS? '{cmd}'", 1)

            return self.inst.ReadString(256)  # reads a maximum of 256 bytes

        else:
            # print('simulation mode of the scope, read VBS')
            print("read, command is : " + str(cmd))

            return "0.123456789"

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
            print("read_float, command is : " + str(cmd))

            return 0.123456789

        pass

    def printScreenToPC(self, path=0, var_name=""):
        # 0 is using the default path, with auto naming from the testing condition

        if path == 0.5:
            # this is the sim_mode for testing
            path = "c:\\py_gary\\test_excel\\test_pic.png  rf"
            pure_path = os.path.splitext(path)[0]
            # . or 'space' will be cut and neglect
            # only path left
            print(pure_path)
            # since there are already PNG in below command, no need to add '.png'
            pass

        elif path == 0:
            if var_name == "":
                # save variable name pic in related folder
                # should separate by the date
                # not change file name if not optional input
                file_name_pic = self.excel_s.wave_condition
                pass

            else:
                file_name_pic = var_name
                pass

            if self.sim_inst == 0:
                pure_path = self.excel_s.wave_path + file_name_pic

            else:
                path = self.excel_s.wave_path + file_name_pic
                # replace all the point to p prevent error of the name
                path = path.replace(".", "p")

            pass

        else:
            # 230503 if using external path rather than
            # save for reserve first

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
            full_path = pure_path + ".png"
            self.excel_s.full_path = full_path

            self.inst.StoreHardcopyToFile(
                "PNG", "BCKG,WHITE,AREA,GRIDAREAONLY", full_path
            )

            pass
        else:
            print("simulation mode of the scope, printScreenToPC")
            pass

        pass

    def printScreenToPC_index(self, file_name0="", reset0=0):
        """
        this function include auto indexing for similiar file name scope capture request
        the final name of the capture will be 'file_name0 + index'
        reset is used to reset the counter set in scope object, start from 0
        """

        # every time call this function, counter will add 1, and reset if reset0 is set to 1
        if reset0 == 1:
            # reset the counter and start from 0 again
            self.capture_index = 0
            pass
        else:
            # increase the index of capture counter save in scope
            self.capture_index = self.capture_index + 1
            pass

        new_name = file_name0 + "_" + str(self.capture_index)

        self.printScreenToPC(path=0, var_name=new_name)
        time.sleep(1)

        pass

    def triggerSingle(self):
        """
        this single will clear sweep after single
        which means that the sample of data will only be once
        """

        if self.sim_inst == 1:
            # logging.debug(f': {cmd}')
            # self.inst.WriteString(f"VBS '{cmd}'", 1)
            # self.writeVBS('app.Acquisition.TriggerMode = "Stopped"')
            self.writeVBS('app.Acquisition.TriggerMode = "Single"')
            self.writeVBS("app.ClearSweeps")

            # self.inst.WriteString(f"VBS 'TriggerDetected = app.Acquisition.Acquire({timeout}, false)'", 1)
            # if 0 == self.inst.WaitForOPC():
            #    raise Exception('LecroyActiveDSO writeVBS timeout or fail.')
            pass
        else:
            print("sopce sim mode, trigger single")
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
            # 230619 : here don't need to import time again
            # import time

            interval_ms = 100

            for t in range(0, timeout_ms, interval_ms):
                if self.readVBS(f"return = app.Acquisition.TriggerMode") == "Stopped":
                    return True
                time.sleep(interval_ms / 1000)

            return False

            '''
            to know success trigger or not, check the returned value of waitTriggered
            if true, good, and if false, bad
            '''


            # if 0 == self.inst.WaitForOPC():
            #     raise Exception('LecroyActiveDSO acquire fail.')

            # r = int(self.readVBS(f'return = TriggerDetected'))

            # if r :
            #     return True
            # else :
            #     raise Exception("LE6100A waitTriggered timeout!")
            pass

        else:
            print("sim mode for scope, wait for trigger")
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
            vestr = f"{value_one_grid:e}"
            ep = vestr.rfind("e")

            M = float(vestr[:ep])
            N = float(vestr[ep + 1 :])

            print(f"vestr={vestr}, M={M}, N={N}")
            if M <= 1:
                vM = 1
            elif M <= 2:
                vM = 2
            elif M <= 5:
                vM = 5
            else:
                vM = 10

            return vM * (10**N)
            pass
        else:
            print("sim mode for scope, find fit scale")

    def ch_default_setting(self):
        # setup channel 1 to 8
        temp_dict = {}
        gen_dict = self.set_general

        for i in range(1, 8 + 1):
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

            self.writeVBS(f'app.Acquisition.C{i}.View = {temp_dict["ch_view"]}')
            self.writeVBS(f'app.Acquisition.C{i}.VerScale = {temp_dict["volt_dev"]}')
            self.writeVBS(f'app.Acquisition.C{i}.BandwidthLimit = "{temp_dict["BW"]}"')
            self.writeVBS(
                f'app.Acquisition.C{i}.EnhanceResType = "{temp_dict["filter"]}"'
            )
            # 221205 add normalization index if selection been set
            if self.nor_v_off == 1:
                # normalize_offset = float(temp_dict["v_offset"]) * temp_dict["v_offset_ind"]
                # self.writeVBS(
                #     f'app.Acquisition.C{i}.VerOffset = {normalize_offset}')
                self.find_signal(i)
                # for the indexed offset, can't be write by it's own, need to
                # know the new mean of signal from find sigfnal function
            else:
                self.writeVBS(
                    f'app.Acquisition.C{i}.VerOffset = {temp_dict["v_offset"]}'
                )

            self.writeVBS(
                f'app.Acquisition.C{i}.LabelsText = "{temp_dict["label_name"]}"'
            )
            self.writeVBS(
                f'app.Acquisition.C{i}.LabelsPosition = "{temp_dict["label_position"]}"'
            )
            self.writeVBS(
                f'app.Acquisition.C{i}.ViewLabels = {temp_dict["label_view"]}'
            )
            self.writeVBS(f'app.Acquisition.C{i}.Coupling = "{temp_dict["coupling"]}"')

            print(f"Grace{self.g_string_a[i-1]}_C{i}")

            if self.sim_inst == 0:
                print(f"simulatiuon mode setting the scope channel{i}")

        # setting of trigger
        self.writeVBS(f'app.Acquisition.TriggerMode = "{gen_dict["trigger_mode"]}"')
        self.writeVBS(
            f'app.Acquisition.Trigger.Source = "{gen_dict["trigger_source"]}"'
        )
        self.writeVBS(
            f'app.Acquisition.Trigger.{gen_dict["trigger_source"]}.Level = {gen_dict["trigger_level"]}'
        )
        self.writeVBS(
            f'app.Acquisition.Trigger.{gen_dict["trigger_source"]}.Slope = "{gen_dict["trigger_slope"]}"'
        )

        # setting of x-axis
        self.writeVBS(f'app.Acquisition.Horizontal.HorScale = {gen_dict["time_scale"]}')
        self.writeVBS(
            f'app.Acquisition.Horizontal.HorOffset = {gen_dict["time_offset"]}'
        )

        # setting of sample mode
        self.writeVBS(
            f'app.Acquisition.Horizontal.SampleMode = "{gen_dict["sample_mode"]}"'
        )

        if self.sim_inst == 0:
            print("setting up general and trigger setting in scope @ simulation mode")

        pass

    def label_chg(self, ch=0, label_name=None, label_position=None, label_view=None):
        if int(ch) != 0:
            ch = str(int(ch))

            if label_name != None:
                self.writeVBS(f'app.Acquisition.C{ch}.LabelsText = "{str(label_name)}"')

            if label_name != None:
                self.writeVBS(
                    f'app.Acquisition.C{ch}.LabelsPosition = "{str(label_position)}"'
                )

            if label_name != None:
                self.writeVBS(f"app.Acquisition.C{ch}.ViewLabels = {str(label_view)}")

        pass

    def inst_name(self):
        # get the insturment name
        self.cmd_str_name = ""
        if self.sim_inst == 1:
            self.cmd_str_name = "*IDN?"
            self.in_name = self.inst_obj.query(self.cmd_str_name)
            time.sleep(0.05)

            pass
        else:
            # for the simulatiom mode of change output
            print("check the instrument name, sim mode ")
            print(str(self.cmd_str_name))
            self.in_name = "scope is in sim mode"

            pass

        return self.in_name

    def scope_initial(self, setup_index):
        """
        socpe initialize function, plan to add loading setting file in scope in the future
        """
        # this function used to choose different setup for different index

        # should recall general setup in the scope first to makesure the items not update in the setup selection be correct

        if setup_index == 0:
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c2 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c3 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c4 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c5 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c6 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c7 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c8 = {
                "ch_view": "TRUE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }

            # setting of general
            self.set_general = {
                "trigger_mode": "Auto",
                "trigger_source": "C3",
                "trigger_level": "-3.2",
                "trigger_slope": "Positive",
                "time_scale": "0.0001",
                "time_offset": "-0.0004",
                "sample_mode": "RealTime",
                "fixed_sample_rate": "1.25GS/s",
            }

            # setting of measurement
            self.p1 = {"param": "max", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "min", "source": "C2", "view": "FALSE"}
            self.p3 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p4 = {"param": "mean", "source": "C4", "view": "FALSE"}
            self.p5 = {"param": "max", "source": "C5", "view": "TRUE"}
            self.p6 = {"param": "min", "source": "C6", "view": "FALSE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "mean", "source": "C6", "view": "FALSE"}
            self.p9 = {"param": "max", "source": "C5", "view": "TRUE"}
            self.p10 = {"param": "min", "source": "C4", "view": "FALSE"}
            self.p11 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "FALSE"}

        elif setup_index == "ripple_50374":
            # index 0 only for functional test
            # the setting for ripple verification
            self.ch_c1 = {
                "ch_view": "TRUE",
                "volt_dev": "0.02",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": -3.3,
                "label_name": "AVDD",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c2 = {
                "ch_view": "TRUE",
                "volt_dev": "0.02",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 3.3,
                "label_name": "OVSS",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c3 = {
                "ch_view": "TRUE",
                "volt_dev": "0.2",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 3.5,
                "label_name": "VON",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c4 = {
                "ch_view": "TRUE",
                "volt_dev": "1",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": -3,
                "label_name": "Vin",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c5 = {
                "ch_view": "TRUE",
                "volt_dev": "0.02",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": -0.06,
                "label_name": "I_load",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC50",
                "v_offset_ind": 1,
            }
            self.ch_c6 = {
                "ch_view": "TRUE",
                "volt_dev": "0.02",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": -3.3,
                "label_name": "OVDD",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c7 = {
                "ch_view": "TRUE",
                "volt_dev": "0.2",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": -3.5,
                "label_name": "VOP",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }
            self.ch_c8 = {
                "ch_view": "FALSE",
                "volt_dev": "0.5",
                "BW": "20MHz",
                "filter": "2bits",
                "v_offset": 0,
                "label_name": "name",
                "label_position": 0,
                "label_view": "TRUE",
                "coupling": "DC1M",
                "v_offset_ind": 1,
            }

            # setting of general
            self.set_general = {
                "trigger_mode": "Auto",
                "trigger_source": "C3",
                "trigger_level": "-3.2",
                "trigger_slope": "Positive",
                "time_scale": "0.0001",
                "time_offset": "-0.0004",
                "sample_mode": "RealTime",
                "fixed_sample_rate": "1.25GS/s",
            }

            # setting of measurement
            self.p1 = {"param": "pkpk", "source": "C1", "view": "TRUE"}
            self.p2 = {"param": "pkpk", "source": "C2", "view": "TRUE"}
            self.p3 = {"param": "pkpk", "source": "C6", "view": "TRUE"}
            self.p4 = {"param": "max", "source": "C3", "view": "TRUE"}
            self.p5 = {"param": "min", "source": "C7", "view": "TRUE"}
            self.p6 = {"param": "mean", "source": "C5", "view": "TRUE"}
            self.p7 = {"param": "pkpk", "source": "C7", "view": "TRUE"}
            self.p8 = {"param": "pkpk", "source": "C3", "view": "TRUE"}
            self.p9 = {"param": "mean", "source": "C3", "view": "TRUE"}
            self.p10 = {"param": "mean", "source": "C4", "view": "TRUE"}
            self.p11 = {"param": "mean", "source": "C6", "view": "TRUE"}
            self.p12 = {"param": "mean", "source": "C2", "view": "TRUE"}

        else:
            # for all the other index, look for setting from the setting object
            self.config_loaded(setup_index)

        # this dictionary is used to assign related measurement
        self.mea_set = {
            "P1": self.p1,
            "P2": self.p2,
            "P3": self.p3,
            "P4": self.p4,
            "P5": self.p5,
            "P6": self.p6,
            "P7": self.p7,
            "P8": self.p8,
            "P9": self.p9,
            "P10": self.p10,
            "P11": self.p11,
            "P12": self.p12,
        }

        # call the 'ch_default_setting' for each channel setting and 'scope_initial'

        # 221213: change the mode to 'Auto' before start giving command, prevent error
        self.trigger_adj(mode="Auto")

        # to speed up the adjustment of scope, need to change sample rate
        self.writeVBS(f'app.Acquisition.Horizontal.Maximize = "FixedSampleRate"')
        self.writeVBS(f'app.Acquisition.Horizontal.SampleRate = "2.5MS/s"')

        self.ch_default_setting()
        self.mea_default_setup()

        self.writeVBS(f'app.Acquisition.Horizontal.Maximize = "FixedSampleRate"')
        self.writeVBS(
            f'app.Acquisition.Horizontal.SampleRate = "{self.set_general["fixed_sample_rate"]}"'
        )

    def config_loaded(self, setup_index):
        # loaded the configuration from external setting object
        self.sc_config.setting_mapping(setup_index)
        # index 0 only for functional test
        # the setting for ripple verification
        self.ch_c1 = self.sc_config.ch_c1
        self.ch_c2 = self.sc_config.ch_c2
        self.ch_c3 = self.sc_config.ch_c3
        self.ch_c4 = self.sc_config.ch_c4
        self.ch_c5 = self.sc_config.ch_c5
        self.ch_c6 = self.sc_config.ch_c6
        self.ch_c7 = self.sc_config.ch_c7
        self.ch_c8 = self.sc_config.ch_c8

        # setting of general
        self.set_general = self.sc_config.set_general
        # setting of measurement
        self.p1 = self.sc_config.p1
        self.p2 = self.sc_config.p2
        self.p3 = self.sc_config.p3
        self.p4 = self.sc_config.p4
        self.p5 = self.sc_config.p5
        self.p6 = self.sc_config.p6
        self.p7 = self.sc_config.p7
        self.p8 = self.sc_config.p8
        self.p9 = self.sc_config.p9
        self.p10 = self.sc_config.p10
        self.p11 = self.sc_config.p11
        self.p12 = self.sc_config.p12

        pass

    def trigger_adj(
        self, mode=None, source=None, level=None, slope=None, clear_sweep=0
    ):
        # this function is used to adjust the trigger function
        """
        choose which trigger to adjust: \n
        mode: 'Auto', 'Normal', 'Single', 'Stopped' \n
        source: 'Cx' \n
        level: 'number' \n
        slope: 'Positive', 'Negative' \n
        """

        if mode != None:
            # add selection prevent re-send the same command causing error
            mode_temp = self.readVBS(f"return = app.Acquisition.TriggerMode")
            if mode_temp != mode:
                if mode == "Single" and clear_sweep == 1:
                    # clear sweep before single
                    self.writeVBS("app.ClearSweeps")
                    time.sleep(0.1)
                self.writeVBS(f'app.Acquisition.TriggerMode = "{mode}"')
                # also change the setting after setting updated
                self.set_general["trigger_mode"] = mode
                pass
            pass
        if source != None:
            self.writeVBS(f'app.Acquisition.Trigger.Source = "{source}"')
            # also change the setting after setting updated
            self.set_general["trigger_source"] = source
            pass
        if level != None and source == None:
            self.writeVBS(
                f'app.Acquisition.Trigger.{self.set_general["trigger_source"]}.Level = {level}'
            )
            # also change the setting after setting updated
            self.set_general["trigger_source"] = level
            pass
        elif level != None and source != None:
            self.writeVBS(f"app.Acquisition.Trigger.{source}.Level = {level}")
            # also change the setting after setting updated
            self.set_general["trigger_level"] = level
            self.set_general["trigger_source"] = source

            pass
        if slope != None and source == None:
            self.writeVBS(
                f'app.Acquisition.Trigger.{self.set_general["trigger_source"]}.Slope = "{slope}"'
            )
            # also change the setting after setting updated
            self.set_general["trigger_source"] = slope
            pass
        elif slope != None and source != None:
            self.writeVBS(f'app.Acquisition.Trigger.{source}.Slope = "{slope}"')
            # also change the setting after setting updated
            self.set_general["trigger_slope"] = slope
            self.set_general["trigger_source"] = source

            pass

        # send current trigger setting for the scope
        print(self.set_general)
        time.sleep(0.5)

        pass

    def Hor_scale_adj(self, div=None, offset=None, sample_rate=0):
        # this program used to adjust the time scale
        """
        to adjust time scale, enter (/div, offset) \n
        sample rate will be default if not setting \n
        sample rate: 1.25GS/s, 2.5GS/s, 5GS/s, 100MS/s, 500MS/s...
        """

        if div != None:
            if float(div) > 0.001:
                # for higher than 1ms, need to set to maximum memory
                self.writeVBS(
                    f'app.Acquisition.Horizontal.Maximize = "SetMaximumMemory"'
                )
                # self.writeVBS(f'app.Acquisition.Horizontal.SampleRate = "2.5GS/s"')

            else:
                # for <= 1ms, using fixed sample rate and set to 2.5G/s

                self.writeVBS(
                    f'app.Acquisition.Horizontal.Maximize = "FixedSampleRate"'
                )
                if sample_rate == 0:
                    self.writeVBS(
                        f'app.Acquisition.Horizontal.SampleRate = "{self.set_general["fixed_sample_rate"]}"'
                    )
                else:
                    self.writeVBS(
                        f'app.Acquisition.Horizontal.SampleRate = "{str(sample_rate)}"'
                    )

            # change the time scale for the scope
            self.writeVBS(f"app.Acquisition.Horizontal.HorScale = {div}")
            # self.set_general['time_scale'] = div
            # this should be default not change

        if offset != None:
            # hcange the offset of zero point in the scope
            self.writeVBS(f"app.Acquisition.Horizontal.HorOffset = {offset}")
            # self.set_general['time_offset'] = offset
            # this should be default not change

        pass

    def single_ch_change(self, ch, ver_scale=None, ver_offset=None, offset_type=0):
        """
        this function change the ver setting of each channel (voltage scale and offset)
        221205: new function added but not done yet, need for adjustment \n
        offset_type set to 1: using normalization offset
        """
        if ver_scale != None:
            # change the ver div
            self.writeVBS(f"app.Acquisition.{ch}.VerScale = {ver_scale}")
            print(f"change {ch} to {ver_scale}")
            pass

        if ver_offset != None and offset_type == 0:
            # change ver offset
            self.writeVBS(f"app.Acquisition.{ch}.VerOffset = {ver_offset}")
            print(f"change {ch} to {ver_offset}")
            pass

        if ver_offset != None and offset_type == 1:
            # the offset input is become index, different way to input the offset setting

            pass
        pass

    def mea_default_setup(self):
        # initial all the measurement channel based on the setting array in scope_initial

        for i in range(1, 12 + 1):
            self.mea_single(
                list(self.mea_set)[i - 1],
                (list(self.mea_set.values())[i - 1])["param"],
                (list(self.mea_set.values())[i - 1])["source"],
                (list(self.mea_set.values())[i - 1])["view"],
            )

            print(list(self.mea_set)[i - 1])
            print((list(self.mea_set.values())[i - 1])["param"])
            print((list(self.mea_set.values())[i - 1])["source"])
            print((list(self.mea_set.values())[i - 1])["view"])

            pass

        pass

    def mea_single(self, mea_ch, parameter, ch=0, view="TRUE"):
        """
        mea_ch: P1-P8, measurement channel \n
        parameter: measurement type \n
        ch default 0, turn off measurment, use 'Cx' as input \n
        view default true, used for initialization \n
        """
        # setup for single channel of measurement

        # turn off measurement if ch set to 0
        if ch != 0:
            # change parameter
            self.writeVBS(f'app.Measure.{mea_ch}.ParamEngine = "{parameter}"')
            # setup mapped scope channel
            self.writeVBS(f'app.Measure.{mea_ch}.Source1 = "{ch}"')

            self.writeVBS(f"app.Measure.{mea_ch}.View = {view}")

        # turn off measurment if ch  is set to 0
        else:
            # visible or not
            self.writeVBS(f"app.Measure.{mea_ch}.View = FALSE")

        pass

    def read_mea(self, mea_ch, m_type, return_float=0):
        """
        mea_ch is measured channel \n
        m_type is last(value), mean, max, exc... \n
        """

        if return_float == 0:
            # choose to return the string for read
            mea_result = self.readVBS(
                f"return = app.Measure.{mea_ch}.{m_type}.Result.Value"
            )
        else:
            # choose to return the float result
            mea_result = self.readVBS_float(
                f"return = app.Measure.{mea_ch}.{m_type}.Result.Value"
            )

        return mea_result

    def capture_full(self, wait_time_s=2, path_t=0, find_level=0, adj_before_save=0):
        """
        this function is the fully scope process, include: clear sweep => Auto =>
        wait time => set to single(trigger) => when trigged stop and capture \n
        find_level: 0 is not to auto find

        """

        self.writeVBS("app.ClearSweeps")
        # 221208 take this line off, since if set to auto mode during auto mode
        # it mat cause extra error
        # self.trigger_adj(mode='Auto')

        if find_level == 1:
            self.find_trig_level()

        if self.sim_inst == 0:
            pass
        else:
            time.sleep(wait_time_s)

        self.trigger_adj(mode="Single")
        self.waitTriggered()

        if adj_before_save == 1:
            # stop to adjust the cursor before capture
            self.excel_s.message_box(
                f"setup the cursor properly and save waveform ",
                "g: stop for cursor",
                auto_exception=1,
                box_type=0,
            )

        self.printScreenToPC(path=path_t)

        pass

    def capture_1st(self, wait_time_s=5, find_level=0, clear_sweep=0):
        """
        stuff need to do before single, contain clear and set to trigger\n
        need to have printScreenToPC after trigger condition is sent\n
        this is the capture full without printScreenToPC
        """

        if self.sim_inst == 0:
            pass
        else:
            # clear sweep and prepare for single
            if clear_sweep == 1:
                self.writeVBS("app.ClearSweeps")
                time.sleep(0.5)
            # time.sleep(wait_time_s)
        if find_level == 0:
            # no need to find level, single directly
            self.trigger_adj(mode="Single")
            # self.waitTriggered()
        else:
            self.find_trig_level()
            self.trigger_adj(mode="Single")
            # self.waitTriggered()
            pass

        pass

    def capture_2nd(self, path_t=0, adj_before_save=0):
        self.waitTriggered()

        if adj_before_save == 1:
            # stop to adjust the cursor before capture
            self.excel_s.message_box(
                f"setup the cursor properly and save waveform ",
                "g: stop for cursor",
                auto_exception=1,
                box_type=0,
            )

        self.printScreenToPC(path=path_t)

        pass

    def cursor_delta(self, abs_val=1, x_y=0, scaling0=1, ch='C1'):
        """
        get the cursor measurement, cursor need to on first\n
        abs default 1 => if not getting abs, set to 0\n
        x_y: 0 -> x and 1 -> y

        230620 => add the channel selection of x and y delta result reading,
        default set to 'C1', can also change when reading different channel delta y
        for the delta of x, all the channel are the same
        """

        if x_y == 0:
            # look for x delta
            temp_res = self.readVBS(
                f"return = app.Cursors.StdCursOf{ch}.DeltaX.Result.Value"
            )
        elif x_y == 1:
            # look for y delta
            temp_res = self.readVBS(
                f"return = app.Cursors.StdCursOf{ch}.DeltaY.Result.Value"
            )

        if abs_val == 1:
            # return the abs
            # 230419: it will have error if no turn on the cursor (error happened here)
            temp_res = float(temp_res)
            temp_res = abs(temp_res) * float(scaling0)

        return temp_res

    def set_cursor(self, x_y0='X', c1_c2=1, view0='true', type0='HorizRel', target0=0):
        '''
        this function used to setup the position of cursor \n
        for x-axis(X), need to know the hol scale settings(in unit: second), \n
        for y-axis(Y), the input is the normalize value, from -4 to 4 \n
        c1_c2(1 or 2), choose which cursor \n
        view type is true or false, default true \n
        type: HorizAbs, HorizRel(two), BothRel(two), VertAbs, VertRel(two)

        '''

        self.writeVBS(f'app.Cursors.View = "{view0}"')
        self.writeVBS(f'app.Cursors.Type = "{type0}"')
        self.writeVBS(f'app.Cursors.{x_y0}Pos{c1_c2} = "{target0}"')

        pass

    def find_trig_level(self, ch=0):
        """
        this program used to change the trigger channel and level and find the
        related level automatically, ch=0 means not change channel
        """
        if ch != 0:
            # change the channel of scope for trigger
            self.trigger_adj(source=f"C{int(ch)}")

        self.writeVBS("app.Acquisition.Trigger.Edge.FindLevel")

        pass

    def find_signal(self, ch=0, variable_index="x"):
        """
        ch is from 1-8
        variable index is the normalization result
        """
        # prevent index error
        print("find signal for g")
        ch = int(ch)
        # change back to the original scale and set the P8 channel back
        ver_scale = float((list(self.sc_config.ch_index.values())[ch - 1])["volt_dev"])

        if (list(self.sc_config.ch_index.values())[ch - 1])["ch_view"] == "TRUE":
            # only active if the channel is turned on

            # first is to change the scale to 10V/div and the offset to 0 (see the signal from -40 to 40V)
            self.single_ch_change(f"C{ch}", 5, 0)
            print("change to 5V/div")

            # change the P8 measurement to related channel and mean, get to know the level of signal
            self.mea_single(mea_ch="P8", parameter="mean", ch=f"C{ch}")

            # clear sweep and collect new mean
            self.writeVBS("app.ClearSweeps")
            time.sleep(0.8)

            new_mean = self.read_mea(mea_ch="P8", m_type="mean", return_float=1)
            if (list(self.sc_config.ch_index.values())[ch - 1])["coupling"] == "AC1M":
                # not going to add the mean if coupling is AC1M
                new_mean = 0
                pass

            new_off_set = (
                -1 * new_mean
                + (list(self.sc_config.ch_index.values())[ch - 1])["v_offset_ind"]
                * ver_scale
            )

            self.single_ch_change(
                ch=f"C{ch}", ver_scale=ver_scale, ver_offset=new_off_set
            )

            # clear sweep and collect new mean
            self.writeVBS("app.ClearSweeps")
            time.sleep(0.8)

            # use the second run for the offset correction
            new_mean = self.read_mea(mea_ch="P8", m_type="mean", return_float=1)
            if (list(self.sc_config.ch_index.values())[ch - 1])["coupling"] == "AC1M":
                # not going to add the mean if coupling is AC1M
                new_mean = 0
                pass

            # only change the second, since this effect the final position
            if variable_index == "x":
                new_off_set = (
                    -1 * new_mean
                    + (list(self.sc_config.ch_index.values())[ch - 1])["v_offset_ind"]
                    * ver_scale
                )
            else:
                new_off_set = -1 * new_mean + float(variable_index) * ver_scale

            """
            issue point: need to use normalization offset index to make sure the waveform is at the same
            place of the scope

            to fix the issue, the function of single_ch_change may also need for adjustment

            """
            self.single_ch_change(
                ch=f"C{ch}", ver_scale=ver_scale, ver_offset=new_off_set
            )

            # need return the measurement change of P8
            self.mea_single(
                mea_ch="P8",
                parameter=self.sc_config.p8["param"],
                ch=self.sc_config.p8["source"],
            )

            pass

        pass

    def ch_view(self, ch, view=1):
        """
        turn off channel to set the view to 0
        """
        # turn on or off the channel

        if view == 1:
            self.writeVBS(f"app.Acquisition.C{ch}.View = TRUE")
        else:
            self.writeVBS(f"app.Acquisition.C{ch}.View = FALSE")

        pass

    def change_label(self, channel0=0, name0=0, position0=0, config=0, view0=1):
        """
        this function is going adjust the label scope with related name and position
        channel
        if channel = 0, pass
        name and position is default 0, config is the default general for reservation
        """

        view0 = int(view0)
        view_str = "TRUE"

        if view0 == 0:
            # disable the lable display
            view_str = "FALSE"

            pass

        else:
            # enalbe the label display

            pass

        if channel0 != 0:
            # operate

            if name0 != 0:
                # only operate the change of name when it's not 0
                self.writeVBS(f'app.Acquisition.C{channel0}.LabelsText = "{name0}"')
            self.writeVBS(f'app.Acquisition.C{channel0}.LabelsPosition = "{position0}"')
            self.writeVBS(f"app.Acquisition.C{channel0}.ViewLabels = {view_str}")

            pass

        else:
            print("channel is set to 0, just pass, no action")

        pass

    def change_setup(self, save0=1, trace0=0, file_name0="", option_function=0):
        """
        this function used to save the scope setup to scope drive
        save-1 => save, save-0 => recall last, save-2 => recall file_name0
        trace default will be C:\ (trace=0)
        file_name is string
        sequence is set to --00000 => can turn off the auto naming
        auto naming default off for Gary's scope setting
        """
        trace_full = "C:\\g_auto_settings\\"

        """
        shoud be able to build the setting files for easier recall, but list down the
        recall name in program or other place
        """

        if trace0 == 0:
            # save the setup in the default trace
            print("default trace selected, doning nothing")

            pass

        else:
            # change to another path for the trace
            trace_full = str(trace0)

            pass

        if save0 == 1:
            # choose to save
            save_s = "Save"

            # update trace
            temp_str = f'app.SaveRecall.Setup.{save_s}SetupFilename = "{trace_full}{file_name0}.lss"'
            print(temp_str)
            self.writeVBS(
                f'app.SaveRecall.Setup.{save_s}SetupFilename = "{trace_full}{file_name0}.lss"'
            )
            pass

        elif save0 == 0:
            # choose to recall
            save_s = "Recall"
            """
            for the recall, read first to get the last saved name
            """
            recall_trace = self.readVBS(
                f"return = app.SaveRecall.Setup.{save_s}SetupFilename.value"
            )

            # update trace
            temp_str = f'app.SaveRecall.Setup.{save_s}SetupFilename = "{recall_trace}"'
            print(temp_str)
            self.writeVBS(
                f'app.SaveRecall.Setup.{save_s}SetupFilename = "{recall_trace}"'
            )
            pass
        elif save0 == 2:
            # choose to recall
            save_s = "Recall"

            # update trace
            temp_str = f'app.SaveRecall.Setup.{save_s}SetupFilename = "{trace_full}{file_name0}.lss"'
            print(temp_str)
            self.writeVBS(
                f'app.SaveRecall.Setup.{save_s}SetupFilename = "{trace_full}{file_name0}.lss"'
            )
            pass



        # run save or recall
        temp_str = f"app.SaveRecall.Setup.Do{save_s}SetupFileDoc2"
        print(temp_str)
        self.writeVBS(f"app.SaveRecall.Setup.Do{save_s}SetupFileDoc2")

        pass


if __name__ == "__main__":
    #  the testing code for this file object
    import parameter_load_obj as par

    excel_t = par.excel_parameter("obj_main")
    sim_scope = 1
    default_path = "C:\\wave_form_raw\\"

    # scope = Scope_LE6100A('GPIB: 5', 3, sim_scope, excel_t)
    scope = Scope_LE6100A(excel0=excel_t)
    # add the path to excel obj
    excel_t.wave_path = default_path
    scope.open_inst()

    test_index = 4
    """
    set 3 to update the channel and others
    set 4 to change the label name
    set 7 to auto waveform saving
    set 5 to save/recall setup
    """

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
        scope.trigger_adj("Auto", "C9", "0.56", "Positive")
        scope.Hor_scale_adj("div_test", "offset_test")
        scope.single_ch_change("C9", "ver_div_test", "ver_off_test")

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
        scope.trigger_adj("Stopped", "C3", "1", "Negative")
        scope.Hor_scale_adj("0.001", "0.00002")
        scope.single_ch_change("C2", "2", "0.25")

        pass

    elif test_index == 3:
        # testing for the measurement setup

        scope.scope_initial("SY8386C_ripple")
        # scope.open_inst()
        temp_name = scope.inst_name()
        print(temp_name)
        # scope.mea_default_setup()

        scope.Hor_scale_adj(0.005, 0)
        scope.Hor_scale_adj(0.0005, 0.0003)
        scope.Hor_scale_adj(0.00025, 0.0001)

        pass

    elif test_index == 4:
        label_pos = 0
        # here is for label name fast change, set index to 4 and run
        # position is in unit, sec
        # maybe plan to add normalize coniguration in future

        # 230426 add label to 0 selection (set to 1 when change to 0, or set to the position you want)
        label_pos_sel = 0
        """
        ms 0.001
        us 0.000001
        """

        if label_pos_sel == 0:
            # list of channel name
            ch_name = {
                "CH1": "EN1",
                "CH2": "3p3V_LDO",
                "CH3": "HSA",
                "CH4": "VDD_IO",
                "CH5": "L_Lx",
                "CH6": "LSCL",
                "CH7": "LSDA",
                "CH8": "VDD_SPD",
            }
        elif label_pos_sel == 1:
            ch_name = {
                "CH1": 0,
                "CH2": 0,
                "CH3": 0,
                "CH4": 0,
                "CH5": 0,
                "CH6": 0,
                "CH7": 0,
                "CH8": 0,
            }
        else:
            ch_name = {
                "CH1": 0,
                "CH2": 0,
                "CH3": 0,
                "CH4": 0,
                "CH5": 0,
                "CH6": 0,
                "CH7": 0,
                "CH8": 0,
            }
            label_pos = label_pos_sel

        # CH1
        scope.change_label(
            channel0=1, name0=ch_name["CH1"], position0=label_pos, config=0, view0=1
        )
        # CH2
        scope.change_label(
            channel0=2, name0=ch_name["CH2"], position0=label_pos, config=0, view0=1
        )
        # CH3
        scope.change_label(
            channel0=3, name0=ch_name["CH3"], position0=label_pos, config=0, view0=1
        )
        # CH4
        scope.change_label(
            channel0=4, name0=ch_name["CH4"], position0=label_pos, config=0, view0=1
        )
        # CH5
        scope.change_label(
            channel0=5, name0=ch_name["CH5"], position0=label_pos, config=0, view0=1
        )
        # CH6
        scope.change_label(
            channel0=6, name0=ch_name["CH6"], position0=label_pos, config=0, view0=1
        )
        # CH7
        scope.change_label(
            channel0=7, name0=ch_name["CH7"], position0=label_pos, config=0, view0=1
        )
        # CH8
        scope.change_label(
            channel0=8, name0=ch_name["CH8"], position0=label_pos, config=0, view0=1
        )

        print("the label setting finished, thanks g")
        pass

    elif test_index == 4.1:
        """
        single change for channel
        """
        ch_num = 8
        name_set = "LX"
        pos = 0

        scope.change_label(
            channel0=ch_num, name0=name_set, position0=pos, config=0, view0=1
        )

        pass

    elif test_index == 5:
        """
        here is plan to add the save and recall setup, to better improve the efficiency of operation
        5 is to save the setup

        save setup can only used for people operation, because same naming cause error need to press enter
        """
        # default_trace = 'C:\\g_auto_settings\\'
        # save_name = '' + '.lss'

        # set to 0 using default trace 'C:\\g_auto_settings\\'
        trace_in = 0
        # BK_glitch
        file_name_in = "NB_bench"
        """
        setup history and list
        compal_seq => compal power on and off sequence
        compal_load_tran => compal load transient
        compal_line_tran => compal line transient
        on_time_check_BK => on time check (for ripple and on time)
        """
        # 1-> save; 0-> recall last; 2-> recall specific
        """
        recall last can be used when change the scope for different capture, but need to recover after nex capature
        """
        save_in = 1
        option_function_in = 0

        scope.change_setup(
            save0=save_in,
            trace0=trace_in,
            file_name0=file_name_in,
            option_function=option_function_in,
        )

        pass

    elif test_index == 6:
        """
        here is plan to add the save and recall setup, to better improve the efficiency of operation
        6 is to recall the setup
        0504: integrated to function 5
        """
        default_trace = "C:\\g_auto_settings\\"
        recall_name = "" + ".lss"

        pass

    elif test_index == 7:
        """
        testing for indexing counter of capture waveform in folder
        """
        x = 0

        while x < 10:
            # index will keep counting when reset0 is 0
            scope.printScreenToPC_index(file_name0="test0", reset0=0)

            x = x + 1
            pass

        # this line reset the index of scope
        scope.printScreenToPC_index(file_name0="test1", reset0=1)

        x = 0
        while x < 5:
            scope.printScreenToPC_index(file_name0="test2", reset0=0)

            x = x + 1
            pass
