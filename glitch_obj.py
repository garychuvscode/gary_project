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

        # object sumulation mode, defaut active (sim mode change to 0)
        self.obj_sim_mode = 0

        # message window for glitch pattern checking
        self.temp_test_mode = 0

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
        self.H_L_excel = self.sh_verification_control.range("B10").value
        self.start_excel = self.sh_verification_control.range("B11").value
        self.step_us_excel = self.sh_verification_control.range("B12").value
        self.count_excel = self.sh_verification_control.range("B13").value
        self.pin_num_excel = int(self.sh_verification_control.range("B14").value)
        self.optional_mode_excel = float(self.sh_verification_control.range("B15").value)
        self.scope_initial_en = int(self.sh_verification_control.range("B16").value)

        # add extra name
        self.excel_ini.extra_file_name = "_glitch"
        self.extra_comments = "glitch"


        # setup the information for each different sheet

        # path need to be assign after every format gen finished
        self.wave_path = self.sh_verification_control.range("C36").value
        self.scope_setting = self.sh_verification_control.range("C39").value
        self.excel_ini.wave_path = self.wave_path
        # the sheet name record for the saving waveform
        self.wave_sheet = self.sh_verification_control.name
        self.excel_ini.wave_sheet = self.wave_sheet

        pass

    def run_verification(self, H_L_pulse=0, start_us0=0, count0=5, step_us0=1, pin_num0=1, scope_set0="970_glitch", optional_mode=0):
        """
        H_L_pulse default setting is low pulse => H-L-H, L programmable \n
        step_us0 = step of each pulse \n
        minimum unit is set to us \n
        pin_num0 = 1 or 2 (default PG1-EN-C4 or PG2-SW-C8)
        optional mode = 1 => define parameter from run_verification input,
        for set to 0, load from the excel sheet
        """

        # reset MCU after each time start
        self.mcu_ini.back_to_initial()
        # pre-test turn on for the instrument
        self.pre_test_inst()
        # load the important parameter
        self.para_loaded()



        if optional_mode == 0 and self.optional_mode_excel == 1 :
            # update the control parameter from excel if optional mode is 0
            H_L_pulse = self.H_L_excel
            start_us0 = self.start_excel
            count0 = self.count_excel
            step_us0 = self.step_us_excel
            pin_num0 = self.pin_num_excel
            scope_set0 = self.scope_setting

            '''
            if self.optional mode == 1 => use the setting from excel and
            run the pattern not follow the raw setting
            '''
            pass
        elif optional_mode == 0 and self.optional_mode_excel == 0 :
            # follow the setting in excel table (just like ripple and transient)
            count0 = self.c_iload
            H_L_pulse = self.H_L_excel
            start_us0 = self.start_excel
            # count0 = self.count_excel
            step_us0 = self.step_us_excel
            pin_num0 = self.pin_num_excel
            scope_set0 = self.scope_setting

            pass



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
            # 230620 trigger falling edge for low pulse
            self.scope_ini.trigger_adj(slope='Negative')
            self.extra_comments = self.extra_comments + f'_{pin_str}_low_pulse'
            pass
        else:
            self.mcu_ini.i_o_change(set_or_clr0=0, pin_num0=pin_num0)
            # pin 1 is EN and pin 2 is SW
            # depends on different IO port send in
            # 230620 trigger rising edge for high pulse
            self.scope_ini.trigger_adj(slope='Positive')
            self.extra_comments = self.extra_comments + f'_{pin_str}_high_pulse'
            pass

        # setting for scope cursor initialize
        # set the X2 to 0 and change X1 is easier
        self.scope_ini.set_cursor(x_y0='X', c1_c2=2, view0='true', type0='HorizRel', target0=0)


        # setting for loader to better see glitch testing result
        self.loader_ini.chg_out2(self.load_curr_EL_LDO, self.excel_ini.loader_ELch, "on")
        self.loader_ini.chg_out2(self.load_curr_VCI_BUCK, self.excel_ini.loader_VCIch, "on")

        # table should be assign when generation of format gen
        self.excel_ini.sh_ref_table.range("B1").value = self.extra_comments

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

                excel_t.message_box("press enater for next pattern", "you're in test mode", auto_exception=self.temp_test_mode)

                # define the length from pulse step
                if optional_mode == 0 and self.optional_mode_excel == 0 :
                    length_us = int(self.excel_ini.sh_format_gen.range(
                            (43 + x_glitch, 7)).value)
                    pass
                else:
                    length_us = start_us0 + (x_glitch) * step_us0
                    pass
                temp_cur_target = length_us * 0.000001
                # change the X2 to related position
                self.scope_ini.set_cursor(target0=temp_cur_target)

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

                # select the related range
                """
                here is to choose related cell for wafeform capture, x is x axis and y is y axis,
                need to input (y, x) for the range input, and x y define is not reverse

                """
                x_index = x_glitch
                y_index = x_vin

                active_range = self.excel_ini.sh_ref_table.range(
                    self.format_start_y + y_index * (2 + self.c_data_mea),
                    self.format_start_x + x_index,
                )
                # (1 + ripple_item) is waveform + ripple item + one current line

                # add the setting of condition in the blank
                active_range.value = f"Vin={v_target}_glitch={length_us}us"

                if self.obj_sim_mode == 0:
                    self.excel_ini.scope_capture(
                        self.excel_ini.sh_ref_table, active_range, default_trace=0.5
                    )
                else:
                    self.excel_ini.scope_capture(
                        self.excel_ini.sh_ref_table, active_range, default_trace=0
                    )
                    pass

                print("check point")

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

    gli_test.run_verification(
        H_L_pulse=1, start_us0=10, count0=20, step_us0=10, pin_num0=1
    )
    time.sleep(2)
    gli_test.run_verification(
        H_L_pulse=0, start_us0=10, count0=20, step_us0=10, pin_num0=1
    )

    excel_t.end_of_file(0)
    print("end of glitch testing program")

    pass
