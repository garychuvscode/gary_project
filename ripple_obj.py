# this file is used to generate the ripple testing file for
# verification object
# need to input the excel obj to reference the format gen sheet and related index
# different pages need to use different index loaded

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


class ripple_test ():

    def __init__(self, excel0, pwr0, met_v0, loader_0, mcu0, src0, met_i0, chamber0):

        # ======== only for object programming
        # testing used temp instrument
        # need to become comment when the OBJ is finished
        import mcu_obj as mcu
        import inst_pkg_d as inst
        import parameter_load_obj as par
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

        # loaded the parameter from the input excel object
        self.c_vin = self.excel_ini.c_ctrl_var1
        self.c_iload = self.excel_ini.c_ctrl_var4
        self.c_data_mea = self.excel_ini.c_data_mea
        self.c_pulse_i2c = self.excel_ini.c_ctrl_var2

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

        self.excel_ini.extra_file_name = '_ripple'



        pass

    def run_verification(self):
        # this function is to run the main item, for all the instrument control and main loop will be in this sub function
        # for the parameter only loaded to the program, no need to call from boject all the time
        # save to local variable every time call the run_verification program

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

        # specific variable for each verification
        loader_ELch = excel_s.loader_ELch
        loader_VCIch = excel_s.loader_VCIch
        wait_time = excel_s.wait_time
        # wait_small = excel_s.wait_small

        en_start_up_check = excel_s.en_start_up_check
        pre_test_en = excel_s.pre_test_en
        relay0_ch = excel_s.relay0_ch
        loader_ELch = excel_s.loader_ELch
        loader_VCIch = excel_s.loader_VCIch
        source_meter_channel = excel_s.source_meter_channel
        en_chamber_mea = excel_s.eff_chamber_en
        inst_auto_selection = excel_s.auto_inst_ctrl
        chamber_default_tset = excel_s.cham_tset_ini
        c_tempature = excel_s.c_tempature
        sw_i2c_select = excel_s.sw_i2c_select
        channel_mode = excel_s.channel_mode
        c_avdd_pulse = excel_s.c_avdd_pulse
        c_pulse = excel_s.c_pulse
        c_i2c = excel_s.c_i2c
        c_i2c_g = excel_s.c_i2c_g
        c_avdd_load = excel_s.c_avdd_load
        wait_small = excel_s.wait_small


        # power supply OV and OC protection
        pwr_s.ov_oc_set(pre_vin_max, pre_imax)

        # power supply channel (channel on setting)
        if pre_test_en == 1:
            pwr_s.chg_out(pre_vin, pre_sup_iout,
                            relay0_ch, 'on')
            print('pre-power on here')
            # turn off the power and load

            load_s.chg_out(0, loader_ELch, 'off')
            load_s.chg_out(0, loader_VCIch, 'off')

            if source_meter_channel == 1 or source_meter_channel == 2:
                load_src_s.load_off()

            print('also turn all load off')

            if en_start_up_check == 1:
                msg_res = win32api.MessageBox(
                    0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(pre_vin) + 'Iin= ' + str(pre_sup_iout))

            if en_chamber_mea == 1:
                # chamber turn on with default setting, using default temperature
                chamber_s.chamber_set(chamber_default_tset)

        # the power will change from initial state directly, not turn off between transition

        # should not need the extra Vin in the change
        # pwr_s.chg_out(vin1_set, pre_sup_iout, relay0_ch, 'on')

        # loader channel and current
        # default off, will be turn on and off based on the loop control

        # load_s.chg_out(iload1_set, loader_ELch, 'off')
        # # load set for EL-power
        # load_s.chg_out(iload2_set, loader_VCIch, 'off')
        # # load set for AVDD

        time.sleep(wait_time)

        # add the while loop outside of SWIRE or I2C loop
        x_temperature = 0

        count_temperature = c_tempature
        if en_chamber_mea == 0:
            # cancel the temperature counter if en_chamber is disable
            count_temperature = 1
        # calculated of c_open_new is below after c_sw_i2c is ready
        # c_open_new = count_temperature * c_sw_i2c
        while x_temperature < count_temperature:
            # update temperature setting every time the loop is start
            tset_now = excel_s.sh_volt_curr_cmd.range(
                (3 + x_temperature, 9)).value
            if en_chamber_mea == 1:
                chamber_s.chamber_set(tset_now)
            else:
                tset_now = 25
                # the temperature without chamber are all assume to be 25C

            # efficiency testing program starts from heere

            # 1st loop is the selection of I2C and SWIRE pulse control

            # selection for the loop control variable
            x_sw_i2c = 0
            c_sw_i2c = 0

            # selection for the SWIRE(1) or I2C(2)
            if sw_i2c_select == 1:
                if channel_mode == 1:
                    # if the setting is avdd only ,need to map the counter to AVDD pulse number
                    c_sw_i2c = c_avdd_pulse
                elif channel_mode == 0 or channel_mode == 2:
                    c_sw_i2c = c_pulse
            elif sw_i2c_select == 2:
                c_sw_i2c = c_i2c
                # i2c group counter setting
                c_i2c_group = c_i2c_g

            if x_temperature == 0:
                c_open_new = count_temperature * c_sw_i2c
                pass


        pass



    # no sheet gen needed since it's been generate by the format gen obj
    # the only thing needed is to

    # since there are more than 1 file for efficiency test, need to call file name reset
    def extra_file_name_setup(self):
        self.excel_ini.extra_file_name = '_ripple'

        pass

    def summary_table ():
        # this sub plan to generate the summary table for each sheet

        pass
