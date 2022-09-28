# this file will create for the general testing
# instrument control is set from main
#


# excel parameter and settings
import parameter_load_obj as par
# for the jump out window
# # also for the jump out window, same group with win32con
import win32api
from win32con import MB_SYSTEMMODAL
# for the delay function
import time
# include for atof function => transfer string to float
import locale as lo


class general_test ():

    def __init__(self, excel0, pwr0, met_v0, loader_0, mcu0, src0, met_i0, chamber0):

        # ======== only for object programming
        # testing used temp instrument
        # need to become comment when the OBJ is finished
        import mcu_obj as mcu
        import inst_pkg_d as inst
        # initial the object and set to simulation mode
        pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
        pwr0.sim_inst = 0
        # initial the object and set to simulation mode
        met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
        met_v0.sim_inst = 0
        loader_0 = inst.chroma_63600(1, 7, 'CCL')
        loader_0.sim_inst = 0
        # mcu is also config as simulation mode
        mcu0 = mcu.MCU_control(0, 3)
        # using the main control book as default
        excel0 = par.excel_parameter('obj_main')
        src0 = inst.Keth_2440(0, 0, 24, 'off', 'CURR', 15)
        src0.sim_inst = 0
        met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
        met_i0.sim_inst = 0
        chamber0 = inst.chamber_su242(25, 10, 'off', -45, 180, 0)
        chamber0.sim_inst = 0
        # ======== only for object programming

        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        # assign the input information to object variable
        self.excel_ini = excel0
        self.pwr_ini = pwr0
        self.loader_ini = loader_0
        self.met_v_ini = met_v0
        self.mcu_ini = mcu0
        self.src_ini = src0
        self.met_i_ini = met_i0
        self.chamber_ini = chamber0
        # self.single_ini = single0

        # # setup extra file name if single verification
        # if self.single_ini == 0:
        #     # this is not single item verififcation
        #     # and this is not the last item (last item)
        #     pass
        # elif self.single_ini == 1:
        #     # it's single, using it' own file name
        #     # item can decide the extra file name is it's the only item
        #     self.excel_ini.extra_file_name = '_SWIRE_pulse'
        #     pass

        self.excel_ini.extra_file_name = '_gen'

    pass

    def sheet_gen(self):

        # copy the rsult sheet to result book
        self.excel_ini.sh_general_test.copy(self.excel_ini.sh_ref)
        # assign the sheet to result book
        self.excel_ini.sh_general_test = self.excel_ini.wb_res.sheets(
            'general')

        # this is both the control and result sheet

        pass

    def run_verification(self):

        # slave object in subprogram
        pwr_s = self.pwr_ini
        load_s = self.loader_ini
        met_v_s = self.met_v_ini
        mcu_s = self.mcu_ini
        excel_s = self.excel_ini
        load_src_s = self.src_ini
        met_i_s = self.met_i_ini
        chamber_s = self.chamber_ini

        # things must have in all run_verification
        pre_vin = excel_s.pre_vin
        pre_sup_iout = excel_s.pre_sup_iout
        pre_imax = excel_s.pre_imax
        pre_vin_max = excel_s.pre_vin_max

        pass
