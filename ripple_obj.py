# this file is used to generate the ripple testing file for
# verification object
# need to input the excel obj to reference the format gen sheet and related index
# different pages need to use different index loaded

# excel parameter and settings
from Scope_LE6100A import Scope_LE6100A
import parameter_load_obj as par
# for the jump out window
# # also for the jump out window, same group with win32con
import win32api
from win32con import MB_SYSTEMMODAL
# for the delay function
import time
# include for atof function => transfer string to float
import locale as lo

import logging as log


class ripple_test ():

    def __init__(self, excel0, pwr0, met_v0, loader_0, mcu0, src0, met_i0, chamber0, scope0):
        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst
            # add the libirary from Geroge
            import Scope_LE6100A as sco
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
            scope0 = sco.Scope_LE6100A('GPIB: 15', 0, 0)
            # ======== only for object programming

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

        pass

    def para_loaded(self):
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
        self.format_start_x = self.excel_ini.format_start_x
        self.format_start_y = self.excel_ini.format_start_y
        self.en_i2c_mode = self.sh_verification_control.range('B10').value
        self.i2c_group = self.sh_verification_control.range('B11').value
        self.c_i2c = self.sh_verification_control.range('B12').value
        self.avdd_current_3ch = self.sh_verification_control.range('B13').value
        self.ch_index = int(self.sh_verification_control.range('B14').value)
        self.c_data_mea = self.excel_ini.c_data_mea

        self.excel_ini.extra_file_name = '_ripple'

        # setup the information for each different sheet

        # path need to be assign after every format gen finished
        self.wave_path = self.sh_verification_control.range('C36').value
        self.excel_ini.wave_path = self.wave_path
        # the sheet name record for the saving waveform
        self.wave_sheet = self.sh_verification_control.name
        self.excel_ini.wave_sheet = self.wave_sheet

        pass

    def pre_test_inst(self):

        # power supply channel (channel on setting)
        if self.excel_ini.pre_test_en == 1:
            self.pwr_ini.chg_out(self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout,
                                 self.excel_ini.relay0_ch, 'on')
            print('pre-power on here')
            # turn off the power and load

            self.loader_ini.chg_out(0, self.excel_ini.loader_ELch, 'off')
            self.loader_ini.chg_out(0, self.excel_ini.loader_VCIch, 'off')

            print('also turn all load off')

            if self.excel_ini.en_start_up_check == 1:
                self.excel_ini.message_box('press enter if hardware configuration is correct',
                                           'Pre-power on for system test under Vin= ' + str(self.excel_ini.pre_vin) + 'Iin= ' + str(self.excel_ini.pre_sup_iout))

        pass

    def run_verification(self):
        '''
        run ripple testing verification
        '''

        self.para_loaded()
        # for the control of temperature, now can be loaded from from the main
        # and separate the sheet name
        # need one variable for the sheet name from main program

        # flexible_naming function is used to add more stuff at the file name
        # call from excel object for the control of flexible file name

        # reset the content of the extra comments every time before the loop start
        extra_comments = ''
        # this used to save the in

        self.pre_test_inst()

        # slave object in subprogram
        pwr_s = self.pwr_ini
        load_s = self.loader_ini
        met_v_s = self.met_v_ini
        mcu_s = self.mcu_ini
        excel_s = self.excel_ini
        load_src_s = self.src_ini
        met_i_s = self.met_i_ini
        chamber_s = self.chamber_ini
        scope_s = self.scope_ini

        # control variable
        if self.en_i2c_mode == 1:
            # I2C mode enable
            c_sw_i2c = self.c_i2c
            # group control for i2c
            c_i2c_group = self.i2c_group
            extra_comments = 'i2c'
            pass

        elif self.en_i2c_mode == 0:
            #  pulse mode enable
            c_sw_i2c = self.c_pulse_i2c
            extra_comments = 'swire_pulse'
            pass

        c_vin = self.c_vin
        c_load_curr = self.c_iload

        if self.ch_index == 0:
            #  EL mode
            extra_comments = extra_comments + '_EL'
            pass
        elif self.ch_index == 1:
            #  VCI mode
            extra_comments = extra_comments + '_VCI'
            pass
        elif self.ch_index == 2:
            #  3-ch mode
            extra_comments = extra_comments + '_3ch'
            pass

        # the loop for pulse and i2c control
        x_sw_i2c = 0
        while x_sw_i2c < c_sw_i2c:

            # assign pulse or i2c command
            if self.en_i2c_mode == 1:
                x_i2c_group = 0

                while x_i2c_group < c_i2c_group:

                    # send the i2c command from MCU

                    # I2C control loop
                    # set up the i2c related data
                    reg_i2c = excel_s.sh_format_gen.range(
                        (43 + c_i2c_group * x_sw_i2c + x_i2c_group, 5)).value
                    data_i2c = excel_s.sh_format_gen.range(
                        (43 + c_i2c_group * x_sw_i2c + x_i2c_group, 6)).value
                    print('register: ' + reg_i2c)
                    print('data: ' + data_i2c)

                    mcu_s.i2c_single_write(reg_i2c, data_i2c)
                    extra_comments = extra_comments + '_' + \
                        str(reg_i2c) + '-' + str(data_i2c)
                    pass

                    # excel_s.sh_ref_table.range('B1')

                # when the command end, need to add the extra information to
                # related sheet for the record testing condition
                # no matter pulse or the i2c command
                # it can be loaded in the blank and not using too long sheet name

                pass

            elif self.en_i2c_mode == 0:
                pulse1 = excel_s.sh_format_gen.range(
                    (43 + x_sw_i2c, 5)).value
                pulse2 = excel_s.sh_format_gen.range(
                    (43 + x_sw_i2c, 6)).value

                mcu_s.pulse_out(pulse1, pulse2)

                extra_comments = extra_comments + '_' + \
                    str(int(pulse1)) + '_' + str(int(pulse2))

                pass

            # table should be assign when generation of format gen
            excel_s.sh_ref_table.range('B1').value = extra_comments

            # the loop for vin
            x_vin = 0
            while x_vin < c_vin:

                # assign vin command on power supply and ideal V
                v_target = excel_s.sh_format_gen.range(
                    (43 + x_vin, 4)).value

                pwr_s.chg_out(v_target, excel_s.pre_imax,
                              excel_s.relay0_ch, 'on')

                pro_status_str = 'Vin:' + str(v_target)
                excel_s.vin_status = str(v_target)
                excel_s.program_status(pro_status_str)

                # the loop for different i load
                x_iload = 0
                while x_iload < c_load_curr:

                    # assign i_load on related channel
                    iload_target = excel_s.sh_format_gen.range(
                        (43 + x_iload, 1)).value

                    pro_status_str = 'setting iload_target current'
                    excel_s.i_el_status = str(iload_target)
                    print(pro_status_str)
                    excel_s.program_status(pro_status_str)

                    # need to be int, not string for self.ch_index
                    if self.ch_index == 0:
                        # EL power settings
                        load_s.chg_out(iload_target, excel_s.loader_ELch, 'on')
                        load_s.chg_out(0, excel_s.loader_VCIch, 'off')

                        pass
                    elif self.ch_index == 1:
                        # VCI power settings
                        load_s.chg_out(
                            iload_target, excel_s.loader_VCIch, 'on')
                        load_s.chg_out(0, excel_s.loader_ELch, 'off')

                        pass
                    elif self.ch_index == 2:
                        # 3-ch power settings
                        load_s.chg_out(iload_target, excel_s.loader_ELch, 'on')

                        # load other target for VCI
                        i_VCI_target = excel_s.sh_format_gen.range('B13')
                        load_s.chg_out(
                            i_VCI_target, excel_s.loader_VCIch, 'on')
                        pass

                    # calibration Vin

                    temp_v = pwr_s.vin_clibrate_singal_met(
                        0, v_target, met_v_s, mcu_s, excel_s)

                    # setup waveform name
                    excel_s.wave_info_update(
                        typ='ripple', v=v_target, i=iload_target)

                    # measure and capture waveform

                    scope_s.printScreenToPC(0)

                    # select teh related range
                    '''
                    here is to choose related cell for the

                    '''
                    y_index = x_iload
                    x_index = x_vin

                    active_range = excel_s.sh_ref_table.range(self.format_start_x + x_index * (2 + self.c_data_mea),
                                                              self.format_start_y + y_index)
                    # (1 + ripple_item) is waveform + ripple item + one current line
                    excel_s.scope_capture(excel_s.sh_ref_table, active_range,
                                          default_trace=0)
                    print('check point')

                    # need to have scope read and scope capture here

                    # need to have all measure data process
                    # OVDD, OVSS and AVDD => adjust needed for run_verification
                    # may not need data latch function?

                    # turn off load and change condition
                    if self.ch_index == 0:
                        # EL power settings
                        load_s.chg_out(0, excel_s.loader_ELch, 'on')
                        load_s.chg_out(0, excel_s.loader_VCIch, 'off')

                        pass
                    elif self.ch_index == 1:
                        # VCI power settings
                        load_s.chg_out(
                            0, excel_s.loader_VCIch, 'on')
                        load_s.chg_out(0, excel_s.loader_ELch, 'off')

                        pass
                    elif self.ch_index == 2:
                        # 3-ch power settings
                        load_s.chg_out(0, excel_s.loader_ELch, 'on')

                    x_iload = x_iload + 1
                    # end of iload loop
                    pass

                x_vin = x_vin + 1
                # the end of vin loop
                pass

            x_sw_i2c = x_sw_i2c + 1
            # the end of i2c or pulse loop
            pass

        # call the file name update after run verification finished
        # and the file name will be ripple if only do ripple or the end is ripple
        # depend on the end of file
        self.extra_file_name_setup()

        # hope to build the summary table after the test is finished
        self.summary_table()
        self.end_of_exp()

        pass

    def end_of_exp(self):
        self.pwr_ini.inst_single_close(self.excel_ini.relay0_ch)

        self.loader_ini.inst_single_close(self.excel_ini.loader_ELch)
        self.loader_ini.inst_single_close(self.excel_ini.loader_VCIch)

        pass

    def general_loop(self):

        # this sub is used to put the general loop format for this kind of sheet and
        # and it can be use to similiar sheet with several stage of loop needed
        # usually the temperature will be at the out side of the loop
        # sequence sohuld be (1)temperature => (2)pulse or i2C command => (3)Vin
        # => (4)loading setting, include Vin calibration =>
        # (5)data latch and waveform capture
        # for temperature control, can also put outside of the loop
        # and this will control by main, adjust from the main program,
        # to reduce the complexity of general loop

        # the general loop will start from pulse

        # control variable
        c_sw_i2c = 5
        c_vin = 5
        c_load_curr = 5

        # the loop for pulse and i2c control
        x_sw_i2c = 0
        while x_sw_i2c < c_sw_i2c:

            # assign pulse or i2c command

            # the loop for vin
            x_vin = 0
            while x_vin < c_vin:

                # assign vin command on power supply and ideal V

                # the loop for different i load
                x_iload = 0
                while x_iload < c_load_curr:

                    # assign i_load on related channel
                    # calibration Vin

                    # measure and capture waveform
                    # turn off load and change condition

                    x_iload = x_iload + 1
                    # end of iload loop
                    pass

                x_vin = x_vin + 1
                # the end of vin loop
                pass

            x_sw_i2c = x_sw_i2c + 1
            # the end of i2c or pulse loop
            pass

        # call the file name update after run verification finished
        # and the file name will be ripple if only do ripple or the end is ripple
        # depend on the end of file
        self.extra_file_name_setup()

        # hope to build the summary table after the test is finished
        self.summary_table()

        pass

    def extra_file_name_setup(self):
        self.excel_ini.extra_file_name = '_ripple'

        pass

    def summary_table(self):
        # this sub plan to generate the summary table for each sheet

        pass


if __name__ == '__main__':
    #  the testing code for this file object

    # ======== only for object programming
    # testing used temp instrument
    # need to become comment when the OBJ is finished
    import mcu_obj as mcu
    import inst_pkg_d as inst
    import Scope_LE6100A as sco
    import parameter_load_obj as par
    # initial the object and set to simulation mode
    excel_t = par.excel_parameter('obj_main')
    pwr_t = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
    pwr_t.sim_inst = 1
    pwr_t.open_inst()
    # initial the object and set to simulation mode
    met_v_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
    met_v_t.sim_inst = 1
    met_v_t.open_inst()
    load_t = inst.chroma_63600(1, 7, 'CCL')
    load_t.sim_inst = 1
    load_t.open_inst()
    met_i_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_i_t.sim_inst = 0
    met_i_t.open_inst()
    src_t = inst.Keth_2440(0, 0, 24, 'off', 'CURR', 15)
    src_t.sim_inst = 0
    src_t.open_inst()
    chamber_t = inst.chamber_su242(25, 10, 'off', -45, 180, 0)
    chamber_t.sim_inst = 0
    chamber_t.open_inst()
    scope_t = sco.Scope_LE6100A('GPIB: 5', 0, 1, excel_t)
    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu_t = mcu.MCU_control(1, 4)
    mcu_t.com_open()

    # for the single test, need to open obj_main first,
    # the real situation is: sheet_ctrl_main_obj will start obj_main first
    # so the file will be open before new excel object benn define

    # using the main control book as default
    # excel_t = par.excel_parameter('obj_main')
    # ======== only for object programming

    # open the result book for saving result
    excel_t.open_result_book()

    # change simulation mode delay (in second)
    excel_t.sim_mode_delay(0.02, 0.01)
    inst.wait_time = 0.01
    inst.wait_samll = 0.01

    import format_gen_obj as form_g

    # for these series of testing, need to add format gen and use format gen

    # and the different verification method can be call below

    version_select = 0

    format_g = form_g.format_gen(excel_t)
    ripple_t = ripple_test(excel_t, pwr_t, met_v_t, load_t,
                           mcu_t, src_t, met_i_t, chamber_t, scope_t)

    if version_select == 0:
        # create one object

        format_g.set_sheet_name('CTRL_sh_ripple')
        format_g.sheet_gen()
        format_g.run_format_gen()

        ripple_t.run_verification()

        format_g.table_return()

        excel_t.end_of_file(0)

    elif version_select == 1:

        excel_t.end_of_file(0)
