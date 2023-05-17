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


        self.setup_index_array = ["NT50970_glitch", "NT50374_glitch" ]

        """
        program setting of this object
        these setting can be change before operation, to save deveop time cost
        update settings directly in the source code
        """
        self.setup_name = self.setup_index_array[0]

        pass

    def para_loaded(self):
        """
        load settings from excel
        """

        pass

    def run_verification(
        self, H_L_pulse=0, start_us0=1, count0=5, step_us0=1, pin_num0=1
    ):
        """
        H_L_pulse default setting is low pulse => H-L-H, L programmable \n
        step_us0 = step of each pulse \n
        minimum unit is set to us \n
        pin_num0 = 1 or 2 (default PG1 or PG2)
        """
        # mapped pint string
        if pin_num0 == 1:
            pin_str = "EN"
            # default state is low, and send high pulse, then back to low
            # other pin don't care (become L)
            single_cell_state1 = "1"
            single_cell_state2 = "0"
        else:
            pin_str = "SW"
            # other pin don't care (become L)
            single_cell_state1 = "2"
            single_cell_state2 = "0"

        # change the initial state of I/O high low
        if H_L_pulse == 0:
            # sending low pulse, no need change initial state
            pass
        else:
            self.mcu_ini.i_o_change(set_or_clr0=0, pin_num0=pin_num0)
            # pin 1 is EN and pin 2 is SW

        # testing of pulse
        x_glitch = 0
        # amount of testing items
        c_glitch = count0
        while x_glitch < c_glitch:
            # define the length from pulse step
            length_us = start_us0 + (x_glitch) * step_us0

            # scope trigger here

            if H_L_pulse == 0:
                # default is low pulse
                self.mcu_ini.g_pulse_out_V2(
                    pulse0=1, duration_ns=1000, en_sw=pin_str, count0=length_us
                )
                pass
            else:
                pattern = (
                    f"'0$1`{single_cell_state1}${length_us}`{single_cell_state2}$1`'"
                )
                full_str = f"mcu.pattern.setupX( {{{pattern}}},1000, 0 )"
                self.mcu_ini.pattern_gen_full_str(cmd_str0=full_str)

                pass

            # scope capture here

            time.sleep(1)
            x_glitch = x_glitch + 1
            pass

        # back to MCU default state
        self.mcu_ini.back_to_initial()

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
        """
        the mapped name doing glitch
        """
        self.excel_ini.extra_file_name = "_glitch"

        pass


if __name__ == "__main__":
    # testing of glitch

    #  the testing code for this file object
    sim_test_set = 0

    import mcu_obj as mcu_msp
    import inst_pkg_d as inst
    import Scope_LE6100A as sco
    import JIGM3 as mcu_g

    # initial the object and set to simulation mode
    excel_t = par.excel_parameter("obj_main")
    pwr_t = inst.LPS_505N(3.7, 0.5, 3, 1, "off")
    pwr_t.sim_inst = sim_test_set
    pwr_t.open_inst()

    # initial the object and set to simulation mode
    met_v_t = inst.Met_34460(0.0001, 30, 0.000001, 2.5, 20)
    met_v_t.sim_inst = sim_test_set
    met_v_t.open_inst()
    load_t = inst.chroma_63600(1, 7, "CCL")
    load_t.sim_inst = sim_test_set
    load_t.open_inst()
    met_i_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_i_t.sim_inst = 0
    met_i_t.open_inst()
    src_t = inst.Keth_2440(0, 0, 24, "off", "CURR", 15)
    src_t.sim_inst = 0
    src_t.open_inst()
    chamber_t = inst.chamber_su242(25, 10, "off", -45, 180, 0)
    chamber_t.sim_inst = 0
    chamber_t.open_inst()
    scope_t = sco.Scope_LE6100A(excel0=excel_t)
    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu_t = mcu_g.JIGM3(sim_mcu0=1, com_addr0=0)
    mcu_t.com_open()

    gli_test = glitch_mea(
        excel_t, pwr_t, met_v_t, load_t, mcu_t, src_t, met_i_t, chamber_t, scope_t
    )

    gli_test.run_verification(
        H_L_pulse=1, start_us0=10, count0=5, step_us0=5, pin_num0=1
    )
    time.sleep(2)
    gli_test.run_verification(
        H_L_pulse=0, start_us0=10, count0=10, step_us0=10, pin_num0=1
    )

    pass
