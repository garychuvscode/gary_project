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
#  220829: for the new structure, using object to define each function
import xlwings as xw


class iq_scan:

    # this class is used to measure IQ from the DUT, based on the I/O setting and different Vin
    # measure the IQ

    def __init__(self, wb_res, pwr, pwr_ch, met_i):
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

        self.control_book_trace = 'c:\\py_gary\\test_excel\\IQ_scan_ctrl.xlsm'
        # no place to load the trace from excel or program, define by default
        # every verification function have independent trace setting
        # no result book trace needed since it's been define by the master excel

        # open sub function control workbook
        self.wb = xw.Book(self.control_book_trace)

        pass

    def sheet_gen(self):
        # this function is a must have function to generate the related excel for this verification item
        # this sub must include:
        # 1. loading the parameter needed for the verification, control loop, instrument or others
        # 2. generate the result sheet in the result book, and setup the format
        # 3. if plot is needed for this verification, need to integrated the plot in the excel file and call from here
        # 4. not a new file but an add on sheet to the result workbook

        pass

    def table_plot(self):
        # this function need to build the plot needed for this verification
        # include the VBA function inside the excel

        pass

    def run_verification(slef):
        #  this function is to run the main item, for all the instrument control and main loop will be in this sub function

        pass
