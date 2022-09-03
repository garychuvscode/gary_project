# 220903: for the new structure, using object to define each function
# SWIRE scan object

# excel parameter and settings
from inst_pkg_d import chroma_63600
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

    def __init__(self, excel0, pwr0, pwr_ch0, met_v0, loader_0, mcu0, single0):

        # ======== only for object programming
        # testing used temp instrument
        # need to become comment when the OBJ is finished
        import mcu_obj as mcu
        import inst_pkg_d as inst
        # initial the object and set to simulation mode
        pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
        pwr.sim_inst = 0
        # initial the object and set to simulation mode
        met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
        met_v0.sim_inst = 0
        loader_0 = inst.chroma_63600(1, 7, 'CCL')
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
        self.pwr_ch_ini = pwr_ch0
        self.loader_ini = loader_0
        self.met_v_ini = met_v0
        self.mcu_ini = mcu0
        self.single_ini = single0

        # setup extra file name if single verification
        if self.single_ini == 0:
            # this is not single item verififcation
            # and this is not the last item (last item)
            pass
        elif self.single_ini == 1:
            # it's single, using it' own file name
            # item can decide the extra file name is it's the only item
            self.excel_ini.extra_file_name = '_SWIRE_pulse'
            pass

        pass

    def sheet_gen(self):
        # this function is a must have function to generate the related excel for this verification item
        # this sub must include:
        # 2. generate the result sheet in the result book, and setup the format
        # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
        # 4. not a new file but an add on sheet to the result workbook

        # copy the rsult sheet to result book
        self.excel_ini.sh_iq_scan.copy(self.excel_ini.sh_ref)
        # assign the sheet to result book
        self.excel_ini.sh_iq_scan = self.excel_ini.wb_res.sheets('SWIRE_scan')

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
        #  this function is to run the main item, for all the instrument control and main loop will be in this sub function
        self.pwr_ini.chg_out(self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout,
                             self.excel_ini.pwr_act_ch, 'on')
        print('pre-power on here')

        if self.excel_ini.en_start_up_check == 1:
            print('window jump out')
            self.msg_res = win32api.MessageBox(
                0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(self.excel_ini.pre_vin) + 'Iin= ' + str(self.excel_ini.pre_sup_iout))

        print('pre-power on state finished and ready for next')
        time.sleep(self.excel_ini.wait_time)

        # the power will change from initial state directly, not turn off between transition
        self.pwr_ini.chg_out(sh.vin1_set, sh.pre_sup_iout, sh.pwr_ch_set, 'on')

        # loader channel and current
        # default off, will be turn on and off based on the loop control

        self.loader_ini.chg_out(sh.iself.loader_ini_set, sh.loa_ch_set, 'off')
        # load set for EL-power
        self.loader_ini.chg_out(sh.iload2_set, sh.loa_ch2_set, 'off')
        # load set for AVDD

        time.sleep(self.excel_ini.wait_time)
        # AVDD measurement will be independant case, not in the loop
        # keep the loop simple and periodic

        # == AVDD measurement

        # call vin calibration
        vin_clibrate_singal_met(vin_ch, sh.pre_vin)
        # after vin calibration, the v_res_temp will be the last vin value
        # assign to related record
        sh.sh_org_tab.range((10, 9)).value = lo.atof(v_res_temp)

        v_res_temp = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_time)
        sh.sh_org_tab.range((10, 4)).value = lo.atof(v_res_temp)

        self.loader_ini.chg_out(sh.iload2_set, sh.loa_ch2_set, 'on')
        time.sleep(self.excel_ini.wait_time)

        # call vin calibration
        vin_clibrate_singal_met(vin_ch, sh.pre_vin)
        # after vin calibration, the v_res_temp will be the last vin value
        # assign to related record
        sh.sh_org_tab.range((10, 10)).value = lo.atof(v_res_temp)

        v_res_temp = self.met_v_ini.mea_v()
        time.sleep(self.excel_ini.wait_time)
        sh.sh_org_tab.range((10, 6)).value = lo.atof(v_res_temp)
        self.loader_ini.chg_out(sh.iload2_set, sh.loa_ch2_set, 'off')

        # == AVDD measurement end

        x_swire = 0
        # counter for SWIRE pulse amount
        while x_swire < sh.c_swire:

            # update the MCU pulse first and wait small time

            # ideal V decide the channel of relay board
            ideal_v = sh.ideal_v_table(x_swire)
            if ideal_v > 0:
                meter_ch_ctrl = 2
            if ideal_v < 0:
                meter_ch_ctrl = 3
            print(meter_ch_ctrl)

            uart_cmd_str = (chr(5) + mcu_cmd_arry[int(meter_ch_ctrl)])
            print(uart_cmd_str)
            mcu_com.write(uart_cmd_str)
            time.sleep(wait_small)
            # input()
            # finished adjust relay channel

            # MCU update SWIRE pulse

            # SWIRE command for the maximum output voltage of ELVDD and ELVSS
            # SWIRE default status need to be high
            pulse1 = sh.sh_org_tab.range((11 + x_swire, 2)).value
            pulse2 = sh.sh_org_tab.range((11 + x_swire, 8)).value

            uart_cmd_str = chr(mode) + chr(int(pulse1)) + chr(int(pulse2))
            print(uart_cmd_str)
            mcu_com.write(uart_cmd_str)
            print('the pulse is ' + str(pulse1) + ' ' + str(pulse2))
            time.sleep(self.excel_ini.wait_time)
            # input()

            # call vin calibration
            vin_clibrate_singal_met(vin_ch, sh.pre_vin)
            # after vin calibration, the v_res_temp will be the last vin value
            # assign to related record
            sh.sh_org_tab.range((11 + x_swire, 9)).value = lo.atof(v_res_temp)

            time.sleep(self.excel_ini.wait_time)
            # measurement start after the SWIRE pulse is set properly
            v_res_temp = self.met_v_ini.mea_v()
            time.sleep(self.excel_ini.wait_time)
            sh.sh_org_tab.range((11 + x_swire, 4)).value = lo.atof(v_res_temp)

            self.loader_ini.chg_out(
                sh.iself.loader_ini_set, sh.loa_ch_set, 'on')
            time.sleep(self.excel_ini.wait_time)

            # call vin calibration
            vin_clibrate_singal_met(vin_ch, sh.pre_vin)
            # after vin calibration, the v_res_temp will be the last vin value
            # assign to related record
            sh.sh_org_tab.range((11 + x_swire, 10)).value = lo.atof(v_res_temp)

            time.sleep(self.excel_ini.wait_time)
            v_res_temp = self.met_v_ini.mea_v()
            time.sleep(self.excel_ini.wait_time)
            sh.sh_org_tab.range((11 + x_swire, 6)).value = lo.atof(v_res_temp)
            self.loader_ini.chg_out(
                sh.iself.loader_ini_set, sh.loa_ch_set, 'off')

            # save the result after each counter finished
            sh.wb_res.save(sh.result_book_trace)

            x_swire = x_swire + 1

            pass

        mcu_com.write(chr(5) + '00')
        print(chr(5) + '00')
        time.sleep(self.excel_ini.wait_time)

        # turn off load and power supply
        self.pwr_ini.change_V(0)
        # only turn off the power supply channel but not the relay
        self.loader_ini.chg_out(0, sh.loa_ch_set, 'off')
        self.loader_ini.chg_out(0, sh.loa_ch2_set, 'off')

        print('finsihed and goodbye')

        self.pwr_ini.chg_out(0, self.excel_ini.pre_sup_iout,
                             self.excel_ini.pwr_act_ch, 'off')
        print('set the output voltage to 0 but keep the current setting')
        print('')
        # self.pwr_ini.inst_close()
        # since inst_close may turn all the channel, may not be a good command for single function
        print('finsihed and goodbye')

        pass


if __name__ == '__main__':
    #  the testing code for this file object

    # ======== only for object programming
    # testing used temp instrument
    # need to become comment when the OBJ is finished
    import mcu_obj as mcu
    import inst_pkg_d as inst
    # initial the object and set to simulation mode
    pwr = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
    pwr.sim_inst = 0
    # initial the object and set to simulation mode
    met_v = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_v.sim_inst = 0
    # mcu is also config as simulation mode
    # COM address of Gary_SONY is 3
    mcu0 = mcu.MCU_control(0, 3)

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

    # and the different verification method can be call below

    # single testing is usuall set single to 1 and one test
    # create one file
    sw_test = sw_scan(excel1, pwr, 3, met_v, mcu0, 1)

    # generate(or copy) the needed sheet to the result book
    sw_test.sheet_gen()

    # start the testing
    sw_test.run_verification()

    # remember that this is only call by main, not by  object
    excel1.end_of_test(0)

    print('end of the SWIRE scan object testing program')
