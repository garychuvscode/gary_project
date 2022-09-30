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

        self.excel_ini.extra_file_name = '_general'

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

        en_start_up_check = excel_s.en_start_up_check
        pre_test_en = excel_s.pre_test_en
        relay0_ch = excel_s.relay0_ch


        chamber_default_tset = excel_s.cham_tset_ini

        # general testing parameter
        en_chamber_mea = excel_s.gen_chamber_en
        gen_loader_en = excel_s.gen_loader_en
        gen_met_i_en = excel_s.gen_met_i_en
        gen_volt_ch_amount = excel_s.gen_volt_ch_amount
        gen_pulse_i2x_en = excel_s.gen_pulse_i2x_en
        gen_loader_ch_amount = excel_s.gen_loader_ch_amount
        gen_pwr_ch_amount = excel_s.gen_pwr_ch_amount
        gen_pwr_i_set = excel_s.gen_pwr_i_set
        gen_col_amount = excel_s.gen_col_amount



        # power supply OV and OC protection
        pwr_s.ov_oc_set(pre_vin_max, pre_imax)

        # power supply channel (channel on setting)
        if pre_test_en == 1:
            pwr_s.chg_out(pre_vin, pre_sup_iout,
                            relay0_ch, 'on')
            print('pre-power on here')
            # turn off the power and load

            print('also turn all load off')

            if en_start_up_check == 1:
                excel_s.message_box('press enter if hardware configuration is correct',
                                    'Pre-power on for system test under Vin= ' + str(pre_vin) + 'Iin= ' + str(pre_sup_iout))
                # msg_res = win32api.MessageBox(
                #     0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(pre_vin) + 'Iin= ' + str(pre_sup_iout))

            if en_chamber_mea == 1:
                    # chamber turn on with default setting, using default temperature
                    chamber_s.chamber_set(chamber_default_tset)





    def set_sheet_name(self, ctrl_sheet_name0):

        # assign the related sheet of each format gen
        self.ctrl_sheet_name = ctrl_sheet_name0

        # sh_format_gen is the sheet can be access from other object
        # load the setting value for instrument
        self.excel_ini.sh_format_gen = self.excel_ini.wb.sheets(
            str(self.ctrl_sheet_name))
        self.sh_format_gen = self.excel_ini.wb.sheets(
            str(self.ctrl_sheet_name))

        # also include the new sheet setting from each different sheet

        # loading the control values
        # 220926: index of counter need to passed to the excel, so other object or instrument
        # is able to reference

        self.c_row_item = self.sh_format_gen.range('C31').value
        self.c_column_item = self.sh_format_gen.range('C32').value
        # c_data_mea is data count
        self.c_data_mea = self.sh_format_gen.range('C33').value
        self.c_ctrl_var1 = self.sh_format_gen.range('D40').value
        self.c_ctrl_var2 = self.sh_format_gen.range('E40').value
        self.c_ctrl_var4 = self.sh_format_gen.range('G40').value

        # also need to assign the variable to the excel obj
        self.excel_ini.c_row_item = self.c_row_item
        self.excel_ini.c_column_item = self.c_column_item
        self.excel_ini.c_data_mea = self.c_data_mea
        self.excel_ini.c_ctrl_var1 = self.c_ctrl_var1
        self.excel_ini.c_ctrl_var2 = self.c_ctrl_var2
        self.excel_ini.c_ctrl_var4 = self.c_ctrl_var4

        self.item_str = self.sh_format_gen.range('C28').value
        self.row_str = self.sh_format_gen.range('C29').value
        self.col_str = self.sh_format_gen.range('C30').value
        self.extra_str = self.sh_format_gen.range('C34').value
        self.color_default = self.sh_format_gen.range('C37').color
        self.color_target = self.sh_format_gen.range('C38').color

        self.target_width = self.sh_format_gen.range('I2').column_width
        self.target_heigh = self.sh_format_gen.range('I2').row_height
        self.default_width = self.sh_format_gen.range('J5').column_width
        self.default_heigh = self.sh_format_gen.range('J5').row_height
        # default height and width is used to prevent shape change of the table
        # target height and width is used to save the waveform capture from scope

        # start to adjust the the format based on the input settings

        self.new_sheet_name = str(self.sh_format_gen.range('C35').value)

        print('sheet name ready')
        self.sheet_name_ready = 1
        pass





        pass

    def update_inst_settings ():



        pass
