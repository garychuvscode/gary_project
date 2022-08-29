# build from 220829
# this file is to setup an object contain all the general used parameter or the general used
# sub program, and this object can be inheritance by the other verification object
# this may help with the generation of general used data and help for
# the decouple of entire system



#  not only one class for all the stuff, separate for few class definition
class MCU_control ():


    def __init__(self, sim_mcu0, com_addr0) :
        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        self.sim_mcu = sim_mcu0
        self.com_addr = com_addr0

        # because MCU will be separate with GPIB for implementation and test
        self.mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
        # array mpaaing for the relay control
        self.meter_ch_ctrl = 0

        # different mode used in the operation
        self.mcu_mode_swire = 1
        self.mcu_mode_sw_en = 3
        self.mcu_mode_I2C = 4
        self.mcu_mode_8_bit_IO = 5
        self.mcu_mode_pat_gen_py = 6
        self.mcu_mode_pat_gen_encode = 7
        self.mcu_mode_pat_gen_direct = 8

        # MCU mapping for different mode control in 2553
        # MCU mapping for RA_GPIO control is mode 5
        # both mode 1 and 2 should defined as dual SWIRE, need to send two pulse command at one time
        # need to be build in the control sheet

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




    pass