"""
this is the glitch testing program, should include pwr, loader
(for different Vin and loading)
mainly MCU(JIGM3 with pattern generator), scope, excel; other may be optional
"""


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


class glitch_mea:
    # this class is used to measure glitch from the DUT
    # send different glitch to test the deglitch of en pin

    def __init__(
        self, excel0, pwr0, met_v0, loader_0, mcu0, src0, met_i0, chamber0, scope0
    ):
        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import JIGM3 as mcu
            import inst_pkg_d as inst

            # add the libirary from Geroge
            import Scope_LE6100A as sco

            # initial the object and set to simulation mode
            pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, "off")
            pwr0.sim_inst = 0
            # initial the object and set to simulation mode
            met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
            met_v0.sim_inst = 0
            loader_0 = inst.chroma_63600(1, 7, "CCL")
            loader_0.sim_inst = 0
            # mcu is also config as simulation mode
            path = mcu.JIGM3.listdevices()
            mcu0 = mcu.JIGM3(devpath=path[0], sim_mcu0=1)
            # set simulation mode or normal mode
            mcu0.sim_mcu = 1
            mcu0.com_open()
            # using the main control book as default
            excel0 = par.excel_parameter("obj_main")
            src0 = inst.Keth_2440(0, 0, 24, "off", "CURR", 15)
            src0.sim_inst = 0
            met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
            met_i0.sim_inst = 0
            chamber0 = inst.chamber_su242(25, 10, "off", -45, 180, 0)
            chamber0.sim_inst = 0
            scope0 = sco.Scope_LE6100A("GPIB: 15", 0, 0)
            # ======== only for object programming

            pass

        # assign the input information to object variable
        self.excel_ini = excel0
        self.pwr_ini = pwr0
        self.loader_ini = loader_0
        self.met_v_ini = met_v0
        self.mcu_ini = mcu0
        self.src_ini = src0
        self.met_i_ini = met_i0
        self.chamber_ini = chamber0
        self.scope_ini = scope0
        # self.single_ini = single0

        """

        """

        pass

    def para_loaded(self):
        """
        load settings from excel
        """

        pass

    def run_verification(self, pmic_buck0=0):
        pass

    def end_of_exp(self):
        self.pwr_ini.inst_single_close(self.excel_ini.relay0_ch)

        self.loader_ini.inst_single_close(self.excel_ini.loader_ELch)
        self.loader_ini.inst_single_close(self.excel_ini.loader_VCIch)

        # also return MCU settings
        self.mcu_ini.back_to_initial()

        self.excel_ini.ready_to_off = 1

        pass

    def extra_file_name_setup(self):
        if self.ripple_line_load == 0:
            self.excel_ini.extra_file_name = "_ripple"
        elif self.ripple_line_load == 1:
            self.excel_ini.extra_file_name = "_line_tran"
        elif self.ripple_line_load == 2:
            self.excel_ini.extra_file_name = "_load_tran"
        elif self.ripple_line_load == 6:
            self.excel_ini.extra_file_name = "_inrush_current"
        elif self.ripple_line_load == 7:
            self.excel_ini.extra_file_name = "_pwr_seq_EN=SW"
        elif self.ripple_line_load == 8:
            self.excel_ini.extra_file_name = "_pwr_seq_EN"
        elif self.ripple_line_load == 9:
            self.excel_ini.extra_file_name = "_pwr_seq_SW"
        elif self.ripple_line_load == 10:
            # add the case with EN rising first
            self.excel_ini.extra_file_name = "_pwr_seq_SW(EN)"

        pass

    def end_of_exp(self):
        self.pwr_ini.inst_single_close(self.excel_ini.relay0_ch)

        self.loader_ini.inst_single_close(self.excel_ini.loader_ELch)
        self.loader_ini.inst_single_close(self.excel_ini.loader_VCIch)

        # also return MCU settings
        self.mcu_ini.back_to_initial()

        self.excel_ini.ready_to_off = 1

        pass


if __name__ == "__main__":
    # testing of glitch

    pass
