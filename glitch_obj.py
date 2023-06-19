"""
this is the glitch testing program, should include pwr, loader
(for different Vin and loading)
mainly MCU(JIGM3 with pattern generator), scope, excel; other may be optional
"""

# turn off the formatter
# fmt: off

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
'''
230616 if only testing glitch without record waveform,
set the test mode to 1, bypass the instrument and
jump message box for each capture
'''
temp_test_mode = 1

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
        # self.single_ini = single0f

        # default setting of two channel current (in A)
        self.load_curr_EL_LDO = 0.1
        self.load_curr_VCI_BUCK = 0.5



        self.setup_index_array = ["NT50970_glitch", "NT50374_glitch" ]

        """
        program setting of this object
        these setting can be change before operation, to save deveop time cost
        update settings directly in the source code
        """
        self.setup_name = self.setup_index_array[0]

        pass

    def change_i_load(self, i_EL_LDO=0.1, i_VCI_BUCK=0.5):
        '''
        used to change the setting of loader current for each channel
        EL_LDO = CH1
        VCI_BUCK = CH2
        '''
        self.load_curr_EL_LDO = i_EL_LDO
        self.load_curr_VCI_BUCK = i_VCI_BUCK

        pass

    def para_loaded(self):
        """
        load settings from excel, same format with excel in ripple
        """

        # since format gen will update the mapped parameter for diferent
        # sheet, need to re-call after change the sheet
        # loaded the parameter from the input excel object
        self.c_vin = self.excel_ini.c_ctrl_var1
        self.c_iload = self.excel_ini.c_ctrl_var4
        self.c_data_mea = self.excel_ini.c_data_mea
        self.c_pulse_i2c = self.excel_ini.c_ctrl_var2
        # record the mapped sheet have related command
        self.sh_verification_control = self.excel_ini.sh_format_gen
        # fixed start point of the format gen (waveform element)
        self.format_start_y = self.excel_ini.format_start_y
        self.format_start_x = self.excel_ini.format_start_x
        self.en_i2c_mode = self.sh_verification_control.range("B10").value
        self.i2c_group = self.sh_verification_control.range("B11").value
        self.c_i2c = self.sh_verification_control.range("B12").value
        self.avdd_current_3ch = self.sh_verification_control.range("B13").value
        self.ch_index = int(self.sh_verification_control.range("B14").value)
        self.scope_adj = 1
        # scope adj is to decide show the output channel or not, not to disable the auto scope
        # setting, it's different, scope initial setting is from sheet (scope_initial_en)

        if self.ch_index > 2:
            # extra channel setting for measurement
            self.ch_index = self.ch_index - 3
            self.scope_adj = 0

        self.ripple_line_load = float(self.sh_verification_control.range("B15").value)
        self.c_data_mea = self.excel_ini.c_data_mea
        self.scope_initial_en = int(self.sh_verification_control.range("B16").value)

        # add extra name
        self.excel_ini.extra_file_name = "_glitch"


        # setup the information for each different sheet

        # path need to be assign after every format gen finished
        self.wave_path = self.sh_verification_control.range("C36").value
        self.scope_setting = self.sh_verification_control.range("C39").value
        self.excel_ini.wave_path = self.wave_path
        # the sheet name record for the saving waveform
        self.wave_sheet = self.sh_verification_control.name
        self.excel_ini.wave_sheet = self.wave_sheet

        pass

    def run_verification(self, H_L_pulse=0, start_us0=1, count0=5, step_us0=1, pin_num0=1, scope_set0="970_glitch"):
        """
        H_L_pulse default setting is low pulse => H-L-H, L programmable \n
        step_us0 = step of each pulse \n
        minimum unit is set to us \n
        pin_num0 = 1 or 2 (default PG1 or PG2)
        """

        # reset MCU after each time start
        self.mcu_ini.back_to_initial()
        # pre-test turn on for the instrument
        self.pre_test_inst()
        # load the important parameter
        self.para_loaded()



        # scope initialization
        if self.scope_initial_en > 0:
            # enable and change to initial
            self.scope_ini.change_setup(save0=0, trace0=0, file_name0=scope_set0)

            """
            this part is for the normalization check, but using the glitch scope
            setting from loading the file
            """

            '''
            0616 note: HOL scale can use 50us/div (consider)
            '''

            # if self.scope_initial_en > 1:
            #     # 221205 added
            #     # turn the offset setting to normalizaiton setting
            #     self.scope_ini.nor_v_off = 1
            #     pass
            # self.scope_ini.scope_initial(self.scope_setting)

        # pwr ovoc setting
        self.pwr_ini.ov_oc_set(self.excel_ini.pre_vin_max, self.excel_ini.pre_imax)

        # mapped pin string

        '''
        config scope to normal mode with different trigger:
        1. set the mode to normal and change the trigger channel to related one
        2. EN_EN2 (C4) => same with pwr_seq in ripple
        3. SW_EN1 (C8) => same with pwr_seq in ripple
        trigger level default set to 1.8V
        '''

        # set the trigger to normal mode
        self.scope_ini.trigger_adj(mode="Normal")


        if pin_num0 == 1:
            pin_str = "EN"
            # default state is low, and send high pulse, then back to low
            # other pin don't care (become L)
            single_cell_state1 = "1"
            single_cell_state2 = "0"
            # set scope to trigger EN_EN2
            self.scope_ini.trigger_adj(source="C4", level=1.8)
        else:
            pin_str = "SW"
            # other pin don't care (become L)
            single_cell_state1 = "2"
            single_cell_state2 = "0"
            # set scope to trigger EN_EN2
            self.scope_ini.trigger_adj(source="C8", level=1.8)

        # change the initial state of I/O high low
        if H_L_pulse == 0:
            # sending low pulse, no need change initial state
            # initial state is high
            pass
        else:
            self.mcu_ini.i_o_change(set_or_clr0=0, pin_num0=pin_num0)
            # pin 1 is EN and pin 2 is SW
            # depends on different IO port send in
            pass


        # setting for loader to better see glitch testing result
        self.loader_ini.chg_out2(self.load_curr_EL_LDO, self.excel_ini.loader_ELch, "on")
        self.loader_ini.chg_out2(self.load_curr_VCI_BUCK, self.excel_ini.loader_VCIch, "on")

        # the loop for vin
        c_vin = self.c_vin
        x_vin = 0
        while x_vin < c_vin:
            # assign vin command on power supply and ideal V


            # power supply setting here
            v_target = self.excel_ini.sh_format_gen.range((43 + x_vin, 4)).value
            self.pwr_ini.chg_out(v_target, self.excel_ini.pre_imax, self.excel_ini.relay0_ch, "on")

            # testing of pulse
            x_glitch = 0
            # amount of testing items
            c_glitch = count0
            while x_glitch < c_glitch:
                # this loop is same with load current level => load current is inner than Vin level

                excel_t.message_box("press enater for next pattern", "you're in test mode", auto_exception=temp_test_mode)

                # define the length from pulse step
                length_us = start_us0 + (x_glitch) * step_us0

                # scope trigger here
                self.scope_ini.capture_1st(clear_sweep=1)

                if H_L_pulse == 0:
                    # default is low pulse
                    self.mcu_ini.g_pulse_out_V2(
                        pulse0=1, duration_ns=1000, en_sw=pin_str, count0=length_us
                    )
                    # since this is low pulse, it can also be achieve by using pulse out function
                    # otherwise, it should be using pattern gen should be a better way
                    pass
                else:
                    pattern = (
                        f"'0$1`{single_cell_state1}${length_us}`{single_cell_state2}$1`'"
                    )
                    full_str = f"mcu.pattern.setupX( {{{pattern}}},1000, 0 )"
                    self.mcu_ini.pattern_gen_full_str(cmd_str0=full_str)

                    pass

                # scope capture here
                self.scope_ini.capture_2nd()



                time.sleep(1)
                x_glitch = x_glitch + 1
                pass

            x_vin = x_vin + 1

        # back to MCU default state
        self.mcu_ini.back_to_initial()
        self.end_of_exp()

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


    def pre_test_inst(self):
        # power supply channel (channel on setting)
        if self.excel_ini.pre_test_en == 1:
            self.pwr_ini.chg_out(
                self.excel_ini.pre_vin,
                self.excel_ini.pre_sup_iout,
                self.excel_ini.relay0_ch,
                "on",
            )
            print("pre-power on here")
            # turn off the power and load

            self.loader_ini.chg_out(0, self.excel_ini.loader_ELch, "off")
            self.loader_ini.chg_out(0, self.excel_ini.loader_VCIch, "off")

            print("also turn all load off")

            if self.excel_ini.en_start_up_check == 1:
                self.excel_ini.message_box(
                    "press enter if hardware configuration is correct",
                    "Pre-power on for system test under Vin= "
                    + str(self.excel_ini.pre_vin)
                    + "Iin= "
                    + str(self.excel_ini.pre_sup_iout),
                )

        pass

if __name__ == "__main__":
    # testing of glitch

    #  the testing code for this file object
    sim_test_set = 0

    import mcu_obj as mcu_msp
    import inst_pkg_d as inst
    import Scope_LE6100A as sco
    import JIGM3 as mcu_g

    # add the excel sheet mapping
    import format_gen_obj as form_g

    import sheet_ctrl_main_obj as sh

    excel_t = par.excel_parameter(str(sh.file_setting))

    excel_t.open_result_book()
    excel_t.excel_save()

    # initial the object and set to simulation mode
    # excel_t = par.excel_parameter("obj_main")
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

    # main_off_line setting can also used to turn off the scope in test mode
    scope_t = sco.Scope_LE6100A(excel0=excel_t, main_off_line0=1)

    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu_t = mcu_g.JIGM3(sim_mcu0=1, com_addr0=0)
    mcu_t.com_open()

    # define the verification item
    gli_test = glitch_mea(
        excel_t, pwr_t, met_v_t, load_t, mcu_t, src_t, met_i_t, chamber_t, scope_t
    )
    # define the format gen needed for glitch testing
    # for the items need waveform capture, need to use format gen to setup
    # related excel sheet
    format_g = form_g.format_gen(excel_t)

    format_g.set_sheet_name('glitch')

    # gli_test.run_verification(
    #     H_L_pulse=1, start_us0=10, count0=20, step_us0=10, pin_num0=1
    # )
    # time.sleep(2)
    gli_test.run_verification(
        H_L_pulse=0, start_us0=10, count0=20, step_us0=10, pin_num0=1
    )

    excel_t.end_of_file(1)
    print("end of glitch testing program")

    pass
