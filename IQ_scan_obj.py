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
import inst_pkg_d as inst
import xlwings as xw
# for the jump out window
import win32api
# also for the jump out window, same group with win32con
import time


import sheet_ctrl_IQ as sh


wait_time = sh.wait_time

# testing used temp instrument
# need to become comment when the OBJ is finished
pwr1 = inst.LPS_505N(0, 0, 1, 7, 'off')


class iq_scan:

    # this class is used to measure IQ from the DUT, based on the I/O setting and different Vin
    # measure the IQ

    def __init__(self, wb_res, pwr, pwr_ch, met_i, mcu):
        # 220831: object format or structure definition
        # 1. parameter: wb_res(the result book), pwr(power supply object), pwr_ch(power supply channel)
        # met_i(current meter)
        # 2.
        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        # assign the input information to object variable
        self.wb_res_main = wb_res
        self.pwr_main = pwr
        self.pwr_ch_main = pwr_ch
        self.met_i_main = met_i
        self.mcu_main = mcu
        # assign the reference sheet generate by the master

        self.control_book_trace = 'c:\\py_gary\\test_excel\\IQ_scan_ctrl.xlsm'
        # no place to load the trace from excel or program, define by default
        # every verification function have independent trace setting
        # no result book trace needed since it's been define by the master excel

        # assign the reference sheet for the result generation in the sub function object
        self.sh_ref = self.wb_res_main.sheets('ref_sh')
        self.sh_ref_condition = self.wb_res_main.sheets('ref_sh2')

        # open sub function control workbook
        self.wb = xw.Book(self.control_book_trace)

        # define the sheets in control book
        self.sh_main = self.wb.sheets('main')
        self.sh_result = self.wb.sheets('IQ_measured')

        pass

    def sheet_gen(self):
        # this function is a must have function to generate the related excel for this verification item
        # this sub must include:
        # 1. loading the parameter needed for the verification, control loop, instrument or others
        # 2. generate the result sheet in the result book, and setup the format
        # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
        # 4. not a new file but an add on sheet to the result workbook

        # copy the sheets to new book
        # for the new sheet generation, located in sheet_gen
        self.sh_main.copy(self.sh_ref_condition)
        self.sh_result.copy(self.sh_ref)

        # assign both sheet to the new sheets in result book
        self.sh_main = self.wb_res_main.sheets('main')
        self.sh_result = self.wb_res_main.sheets('IQ_measured')

        pass

    def table_plot(self):
        # this function need to build the plot needed for this verification
        # include the VBA function inside the excel

        pass

    def run_verification(self):
        #  this function is to run the main item, for all the instrument control and main loop will be in this sub function
        pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.pwr_ch_set, 'on')
        print('pre-power on here')

        msg_res = win32api.MessageBox(
            0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

        time.sleep(wait_time)

        pass
