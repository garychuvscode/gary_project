#  define the structure general class for future work efficiency

'''
here is define as the long comments for python
here will shown the general format for the class definition, used for future example

class 0_class_name :
    # general class description, and how to use the class
    # easy function

    def __init__(self, input_parameter_0) :
        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        pass

    def sheet_gen(self) :
        # this function is a must have function to generate the related excel for this verification item
        # this sub must include:
        # 1. loading the parameter needed for the verification, control loop, instrument or others
        # 2. generate the result sheet in the result book, and setup the format
        # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
        # 4. not a new file but an add on sheet to the result workbook

        pass

    def table_plot(self) :
        # this function need to build the plot needed for this verification
        # include the VBA function inside the excel

        pass

    def run_verification(slef) :
        #  this function is to run the main item, for all the instrument control and main loop will be in this sub function

        pass



'''
# 220829: for the new structure, using object to define each function

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


class iq_scan:

    # this class is used to measure IQ from the DUT, based on the I/O setting and different Vin
    # measure the IQ

    def __init__(self, excel0, pwr0, pwr_ch0, met_i0, mcu0, single0):

        # # ======== only for object programming
        # # testing used temp instrument
        # # need to become comment when the OBJ is finished
        # import mcu_obj as mcu
        # import inst_pkg_d as inst
        # # initial the object and set to simulation mode
        # pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
        # pwr0.sim_inst = 0
        # # initial the object and set to simulation mode
        # met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
        # met_i0.sim_inst = 0
        # # mcu is also config as simulation mode
        # mcu0 = mcu.MCU_control(0, 3)
        # # using the main control book as default
        # excel0 = par.excel_parameter('obj_main')
        # # ======== only for object programming

        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        # assign the input information to object variable
        self.excel_ini = excel0
        self.pwr_ini = pwr0
        self.pwr_ch_ini = pwr_ch0
        self.met_i_ini = met_i0
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
            self.excel_ini.extra_file_name = '_IQ'
            pass
        # 220903
        # extra program name for multi-verification change to excel object
        # elif single == 2:
        #     # this means it's the last test item for the whole test
        #     # marked up what is the testing program number of this test
        #     # easier to check with the manual sheet
        #     self.extra_file_name = '_P' + \
        #         str(self.excel_ini.program_group_index)
        #     pass

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
        self.excel_ini.sh_iq_scan = self.excel_ini.wb_res.sheets('IQ_measured')

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

        # power supply OV and OC protection
        self.pwr_ini.ov_oc_set(self.excel_ini.pre_vin_max,
                               self.excel_ini.pre_imax)

        self.pwr_ini.chg_out(self.excel_ini.pre_vin, self.excel_ini.pre_sup_iout,
                             self.excel_ini.pwr_act_ch, 'on')
        print('pre-power on here')

        if self.excel_ini.en_start_up_check == 1:
            print('window jump out')
            self.msg_res = win32api.MessageBox(
                0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(self.excel_ini.pre_vin) + 'Iin= ' + str(self.excel_ini.pre_sup_iout))

        print('pre-power on state finished and ready for next')
        time.sleep(self.excel_ini.wait_time)

        x_iq = 0
        # counter for SWIRE pulse amount
        while x_iq < self.excel_ini.c_iq:
            # load the Vin command first
            ideal_v = self.excel_ini.sh_iq_scan.range((7, 4 + x_iq)).value

            # update the vin setting for different vin demand
            self.pwr_ini.chg_out(
                ideal_v, self.excel_ini.pre_sup_iout, self.excel_ini.pwr_act_ch, 'on')

            # four different mode in this loop will change
            x_submode = 0
            # initialize the counter for differnt state
            while x_submode < 4:
                # submode + 1 will be the state command for MCU

                self.mcu_ini.pmic_mode(x_submode + 1)
                # change the mode for IQ measurement
                time.sleep(self.excel_ini.wait_time)
                if x_submode == 0:
                    time.sleep(3 * self.excel_ini.wait_time)
                    # extra wait time for the normal mode to shtudown mode transition
                    # because IQ may not stop change so fast, need to double check

                print('input voltage setting is ' + str(ideal_v))
                print('the mode is ' + str(x_submode + 1))
                time.sleep(self.excel_ini.wait_time)

                # this part can not be in the simulation mode,
                # because it need to access the meter object variable
                if x_submode == 0:
                    # for the mode of measure ISD, need to have more wait time for stable and
                    # need to prevend negative result of error result

                    # here used the change of range directly => change variable in inst_pkg
                    range_temp = self.met_i_ini.max_mea_i_ini
                    self.met_i_ini.max_mea_i_ini = self.excel_ini.ISD_range
                    time.sleep(5 * self.excel_ini.wait_time)
                else:
                    self.met_i_ini.max_mea_i_ini = range_temp
                    # because the counter starts from 0, ini value will save to the temp first and return when the counter
                    # is not 0

                # measurement start after the AVDDEN and SWIRE is updated
                time.sleep(self.excel_ini.wait_small)
                v_res_temp = self.met_i_ini.mea_i()

                if x_submode == 0:
                    while float(v_res_temp) < 0:
                        # when there is a negative result from the measurement
                        # we need to re measure to update the result
                        v_res_temp = self.met_i_ini.mea_i()
                        time.sleep(self.excel_ini.wait_small)
                        # it should be already stable after the change of GPIO command
                        # small wait time should be enough

                # when the measurement is finished, update the result to excel table and map to the scaling
                time.sleep(self.excel_ini.wait_small)
                self.excel_ini.sh_iq_scan.range((8 + x_submode, 4 + x_iq)
                                                ).value = lo.atof(v_res_temp) * self.excel_ini.iq_scaling
                # iq_scaling is decide from the result table (unit is optional)
                x_submode = x_submode + 1

            # update the counter for different Vin
            # self.excel_ini.sh_iq_scan.range((6, 4 + x_iq)).value = ideal_v

            x_iq = x_iq + 1
            self.excel_ini.excel_save()
        # save the result after each counter finished
        # self.excel_ini.wb_res.save(self.excel_ini.result_book_trace)
        # 220903: end of test only call by main, because
        # not knowing if this is single or not
        # excel_ini.end_of_test()

        self.pwr_ini.chg_out(0, self.excel_ini.pre_sup_iout,
                             self.excel_ini.pwr_act_ch, 'off')
        print('set the output voltage to 0 but keep the current setting')
        print('Gary is lucky to meet Grace XD')
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
    met_i = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
    met_i.sim_inst = 0
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
    iq_test = iq_scan(excel1, pwr, 3, met_i, mcu0, 1)

    # generate(or copy) the needed sheet to the result book
    iq_test.sheet_gen()

    # start the testing
    iq_test.run_verification()

    # remember that this is only call by main, not by  object
    excel1.end_of_test(0)

    print('end of the IQ object testing program')
