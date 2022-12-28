# 220903: for the new structure, using object to define each function
# SWIRE scan object

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


class sw_scan:

    # this class is used to measure IQ from the DUT, based on the I/O setting and different Vin
    # measure the IQ

    def __init__(self, excel0, pwr0, met_v0, loader_0, mcu0):

        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst
            # add the libirary from Geroge
            # initial the object and set to simulation mode
            pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
            pwr0.sim_inst = 0
            # initial the object and set to simulation mode
            met_v0 = inst.Met_34460(0.0001, 30, 0.000001, 2.5, 21)
            met_v0.sim_inst = 0
            loader_0 = inst.chroma_63600(1, 7, 'CCL')
            loader_0.sim_inst = 0
            # mcu is also config as simulation mode
            mcu0 = mcu.MCU_control(0, 3)
            # using the main control book as default
            excel0 = par.excel_parameter('obj_main')
            # ======== only for object programming

        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        # assign the input information to object variable
        self.excel_ini = excel0
        self.pwr_ini = pwr0
        # self.pwr_ch_ini = pwr_ch0
        self.loader_ini = loader_0
        self.met_v_ini = met_v0
        self.mcu_ini = mcu0
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
        self.excel_ini.extra_file_name = '_SWIRE_pulse'

        pass

    def extra_file_name_setup(self):
        self.excel_ini.extra_file_name = '_SWIRE_pulse'
        pass

    def sheet_gen(self):
        # this function is a must have function to generate the related excel for this verification item
        # this sub must include:
        # 2. generate the result sheet in the result book, and setup the format
        # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
        # 4. not a new file but an add on sheet to the result workbook

        # copy the rsult sheet to result book
        # self.excel_ini.sh_sw_scan.copy(self.excel_ini.sh_ref)
        # # assign the sheet to result book
        # self.excel_ini.sh_sw_scan = self.excel_ini.wb_res.sheets(
        #     self.excel_ini.sh_sw_scan_name)
        # 221209: since .copy will return the cpoied sheet, just assign, no need for name
        self.excel_ini.sh_sw_scan = self.excel_ini.sh_sw_scan.copy(
            self.excel_ini.sh_ref)

        # # copy the sheets to new book
        # # for the new sheet generation, located in sheet_gen
        # self.sh_main.copy(self.sh_ref_condition)
        # self.sh_result.copy(self.sh_ref)

        pass

    def table_plot(self):
        # this function need to build the plot needed for this verification
        # include the VBA function inside the excel

        pass

    def run_verification(self):
        # this function is to run the main item, for all the instrument control and main loop will be in this sub function
        # for the parameter only loaded to the program, no need to call from boject all the time
        # save to local variable every time call the run_verification progra

        # slave object in subprogram
        pwr_s = self.pwr_ini
        load_s = self.loader_ini
        met_v_s = self.met_v_ini
        mcu_s = self.mcu_ini
        excel_s = self.excel_ini

        # things must have in all run_verification
        pre_vin = excel_s.pre_vin
        pre_sup_iout = excel_s.pre_sup_iout
        pre_imax = excel_s.pre_imax
        pre_vin_max = excel_s.pre_vin_max
        pre_test_en = excel_s.pre_test_en

        self.sheet_gen()

        # make sure MCU back to initial
        self.mcu_ini.back_to_initial()

        # sheet needed in the sub
        res_sheet = excel_s.sh_sw_scan

        # specific variable for each verification
        # 220910 cancel pwr_act_ch, and replaced by relay0_ch
        # pwr_act_ch = excel_s.pwr_act_ch
        vin_set = excel_s.vin_set
        iin_set = excel_s.Iin_set
        EL_curr = excel_s.EL_curr
        VCI_curr = excel_s.VCI_curr
        loader_ELch = excel_s.loader_ELch
        loader_VCIch = excel_s.loader_VCIch
        c_swire = excel_s.c_swire
        wait_time = excel_s.wait_time
        relay0_ch = excel_s.relay0_ch
        # wait_small = excel_s.wait_small

        en_start_up_check = excel_s.en_start_up_check

        excel_s.current_item_index = 'sw_scan'
        # move all as much as possible index to the front of sub program and easier to modify
        # and less chance to fail when change index

        # power supply OV and OC protection
        pwr_s.ov_oc_set(pre_vin_max, pre_imax)

        if pre_test_en == 1:
            pwr_s.chg_out(pre_vin, pre_sup_iout, relay0_ch, 'on')
            print('pre-power on here')

        if en_start_up_check == 1:
            print('window jump out')
            excel_s.message_box('press enter if hardware configuration is correct',
                                'Pre-power on for system test under Vin= ' + str(excel_s.pre_vin) + 'Iin= ' + str(excel_s.pre_sup_iout))
            # self.msg_res = win32api.MessageBox(
            #     0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(excel_s.pre_vin) + 'Iin= ' + str(excel_s.pre_sup_iout))

        print('pre-power on state finished and ready for next')
        time.sleep(wait_time)

        # the power will change from initial state directly, not turn off between transition
        pwr_s.chg_out(vin_set, iin_set, relay0_ch, 'on')

        # loader channel and current
        # default off, will be turn on and off based on loop control

        load_s.chg_out(EL_curr, loader_ELch, 'off')
        # load set for EL-power
        load_s.chg_out(VCI_curr, loader_VCIch, 'off')
        # load set for AVDD

        time.sleep(wait_time)
        # AVDD measurement will be independant case, not in the loop
        # keep the loop simple and periodic

        # == AVDD measurement
        meter_ch_ctrl = 1

        # call vin calibration
        v_res_temp = pwr_s.vin_clibrate_singal_met(
            0, vin_set, met_v_s, mcu_s, excel_s)
        # after vin calibration, the v_res_temp will be the last vin value
        # assign to related record
        res_sheet.range((10, 9)).value = lo.atof(v_res_temp)

        mcu_s.relay_ctrl(meter_ch_ctrl)
        v_res_temp = met_v_s.mea_v()
        time.sleep(wait_time)
        res_sheet.range((10, 4)).value = lo.atof(v_res_temp)

        # self.loader_ini.chg_out(excel_s.VCI_curr, excel_s.loader_VCIch, 'on')
        # change to new format of command, only change loader state (fix load current at swire scan)
        self.loader_ini.chg_state_single(loader_VCIch, 'on')
        time.sleep(wait_time)

        # call vin calibration
        v_res_temp = pwr_s.vin_clibrate_singal_met(
            0, vin_set, met_v_s, mcu_s, excel_s)
        # after vin calibration, the v_res_temp will be the last vin value
        # assign to related record
        res_sheet.range((10, 10)).value = lo.atof(v_res_temp)

        mcu_s.relay_ctrl(meter_ch_ctrl)
        v_res_temp = self.met_v_ini.mea_v()
        time.sleep(wait_time)
        res_sheet.range((10, 6)).value = lo.atof(v_res_temp)
        self.loader_ini.chg_state_single(loader_VCIch, 'off')

        # == AVDD measurement end

        x_swire = 0
        # counter for SWIRE pulse amount
        while x_swire < c_swire:

            # update the MCU pulse first and wait small time

            # ideal V decide the channel of relay board
            ideal_v = excel_s.ideal_v_table(x_swire)

            # mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
            # # array mpaaing for the relay control
            # meter_ch_ctrl = 0
            # # meter channel indicator: 0: Vin, 1: AVDD, 2: OVDD, 3: OVSS, 4: VOP, 5: VON

            if ideal_v > 0:
                meter_ch_ctrl = 2
            if ideal_v < 0:
                meter_ch_ctrl = 3
            print(meter_ch_ctrl)

            mcu_s.relay_ctrl(meter_ch_ctrl)
            # input()
            # finished adjust relay channel

            # MCU update SWIRE pulse

            # SWIRE command for the maximum output voltage of ELVDD and ELVSS
            # SWIRE default status need to be high
            pulse1 = res_sheet.range((11 + x_swire, 2)).value
            pulse2 = res_sheet.range((11 + x_swire, 8)).value

            mcu_s.pulse_out(pulse1, pulse2)
            print('the pulse is ' + str(pulse1) + ' ' + str(pulse2))
            time.sleep(wait_time)
            # input()
            # 221202: add one more command send to make sure the pulse is received
            mcu_s.pulse_out(pulse1, pulse2)
            time.sleep(wait_time)

            # call vin calibration
            v_res_temp = pwr_s.vin_clibrate_singal_met(
                0, vin_set, met_v_s, mcu_s, excel_s)
            # after vin calibration, the v_res_temp will be the last vin value
            # assign to related record
            res_sheet.range((11 + x_swire, 9)).value = lo.atof(v_res_temp)

            time.sleep(wait_time)
            # measurement start after the SWIRE pulse is set properly
            v_res_temp = self.met_v_ini.mea_v()
            time.sleep(wait_time)
            res_sheet.range((11 + x_swire, 4)).value = lo.atof(v_res_temp)

            load_s.chg_state_single(loader_ELch, 'on')
            time.sleep(wait_time)

            # call vin calibration
            v_res_temp = pwr_s.vin_clibrate_singal_met(
                0, vin_set, met_v_s, mcu_s, excel_s)
            # after vin calibration, the v_res_temp will be the last vin value
            # assign to related record
            res_sheet.range((11 + x_swire, 10)).value = lo.atof(v_res_temp)

            time.sleep(wait_time)
            v_res_temp = self.met_v_ini.mea_v()
            time.sleep(wait_time)
            res_sheet.range((11 + x_swire, 6)).value = lo.atof(v_res_temp)
            load_s.chg_state_single(loader_ELch, 'off')

            # save the result after each counter finished
            excel_s.excel_save()
            if self.excel_ini.turn_inst_off == 1:
                self.end_of_exp()
                self.excel_ini.excel_save()

            x_swire = x_swire + 1

            pass

        self.end_of_exp()

        pass

    def end_of_exp(self):
        # reset MCU back to default
        self.mcu_ini.back_to_initial()
        time.sleep(self.excel_ini.wait_time)

        # turn off load and power supply
        self.pwr_ini.change_V(0, self.excel_ini.relay0_ch)
        # only turn off the power supply channel but not the relay
        self.loader_ini.chg_out(0, self.excel_ini.loader_ELch, 'off')
        self.loader_ini.chg_out(0, self.excel_ini.loader_VCIch, 'off')

        self.pwr_ini.chg_out(0, self.excel_ini.pre_sup_iout,
                             self.excel_ini.relay0_ch, 'off')
        print('set the output voltage to 0 but keep the current setting')
        print("Grace's one laugh can make me happy one day!")
        time.sleep(self.excel_ini.wait_time)
        # self.pwr_ini.inst_close()
        # since inst_close may turn all the channel, may not be a good command for single function
        self.extra_file_name_setup()
        print('finsihed and goodbye')

        self.excel_ini.ready_to_off = 1

        pass


if __name__ == '__main__':
    #  the testing code for this file object

    # ======== only for object programming
    # testing used temp instrument
    # need to become comment when the OBJ is finished
    import mcu_obj as mcu
    import inst_pkg_d as inst
    # initial the object and set to simulation mode
    pwr_t = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
    pwr_t.sim_inst = 0
    # initial the object and set to simulation mode
    met_v_t = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_v_t.sim_inst = 0
    load_t = inst.chroma_63600(1, 7, 'CCL')
    load_t.sim_inst = 0
    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu_t = mcu.MCU_control(0, 3)

    # for the single test, need to open obj_main first,
    # the real situation is: sheet_ctrl_main_obj will start obj_main first
    # so the file will be open before new excel object benn define

    # using the main control book as default
    excel1 = par.excel_parameter('obj_main')
    # ======== only for object programming

    # open the result book for saving result
    excel1.open_result_book()

    # change simulation mode delay (in second)
    excel1.sim_mode_delay(0.02, 0.01)
    inst.wait_time = 0.01
    inst.wait_samll = 0.01

    # and the different verification method can be call below

    # create one file
    sw_test = sw_scan(excel1, pwr_t, met_v_t, load_t, mcu_t)

    # generate(or copy) the needed sheet to the result book
    sw_test.sheet_gen()

    # start the testing
    sw_test.run_verification()

    # remember that this is only call by main, not by  object
    excel1.end_of_file(0)

    print('end of the SWIRE scan object testing program')
