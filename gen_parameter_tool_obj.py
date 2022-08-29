# build from 220829
# this file is to setup an object contain all the general used parameter or the general used
# sub program, and this object can be inheritance by the other verification object
# this may help with the generation of general used data and help for
# the decouple of entire system

import time
import pyvisa
rm = pyvisa.ResourceManager()

#  not only one class for all the stuff, separate for few class definition


class MCU_control ():

    def __init__(self, sim_mcu0, com_addr0):
        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        self.sim_mcu = sim_mcu0
        self.com_addr = com_addr0

        # because MCU will be separate with GPIB for implementation and test
        self.mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
        # array mpaaing for the relay control
        self.meter_ch_ctrl = 0

        self.pulse1 = 0
        self.pulse2 = 0
        self.mode_set = 4
        # mode sequence: 1-4: (EN, SW) = (0, 0),  (0, 1), (1, 0), (1, 1) => default normal

        # i2c register and data (single byte data); slave address is fixed in the MCU
        # here only support for the register and data change
        self.reg_i2c = ''
        self.data_i2c = ''

        #  UART command string
        self.uart_cmd_str = ''

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

        self.wait_small = 0.2
        #  waiting time duration, can be adjust with MCU C code

        pass

    def mcu_write(self, index):
        # the MCU write used to generate the command string and send the command out
        # but the related control parameter need to define in the main program
        # here is only used to reduce the code of generate the string
        # and the MCU UART sending command

        if index == 'swire':
            self.uart_cmd_str = (chr(self.mcu_mode_swire) +
                                 chr(int(self.pulse1)) + chr(int(self.pulse2)))
            # for the SWIRE mode of 2553, there are 2 pulse send to the MCU and DUT
            # pulse amount is from 1 to 255, not sure if 0 will have error or not yet
            # 20220121
        elif index == 'en_sw':
            self.uart_cmd_str = (chr(self.mcu_mode_sw_en) +
                                 chr(int(self.mode_set)) + chr(1))
            # for the EN SWIRE control mode, need to handle the recover to normal mode (EN, SW) = (1, 1)
            # at the end of application
            # this mode only care about the first data ( 0-4 )
        elif index == 'relay':
            self.uart_cmd_str = (
                chr(self.mcu_mode_8_bit_IO) + self.mcu_cmd_arry[self.meter_ch_ctrl])
            # assign relay to related channel after function called
            # channel index is from golbal variable

        elif index == 'i2c':
            self.uart_cmd_str = (chr(self.mcu_mode_I2C) +
                                 str(self.reg_i2c) + str(self.data_i2c))
            # send mapped i2c command out from MCU

        else:
            print('wrong command on MCU, double check')
            #  if sending the wrong string, gaive the message through terminal

        # print the command going to send before write to MCU, used for debug
        print(self.uart_cmd_str)
        if self.sim_mcu == 1:
            # now is real mode, output the MCU command from COM port
            self.mcu_com.write(self.uart_cmd_str)
        else:
            print('now is sending the MCU command with below string:')
            print(str(index))
            print(self.uart_cmd_str)

        # give some response time for the UART command send and MCU action
        time.sleep(self.wait_small)

    def com_open(self):
        # this function is used to open the com port of the MCU
        # this will be set independentlly in each object
        print('now the COM port is on')
        uart_cmd_str = "COM" + str(int(self.com_addr))
        print(uart_cmd_str)
        if self.sim_mcu == 1:
            self.mcu_com = rm.open_resource(uart_cmd_str)
        else:
            print('open COM port but bypass the real operation')

        pass

    def com_close(self):
        # after the verification is finished, reset all the condition
        # to initial and turn off the communication port
        self.back_to_initial()

        if self.sim_mcu == 1:
            self.mcu_com.close()
        else:
            print('the com port is turn off now')

        pass

    def back_to_initial(self):
        #  this sub program used to set all the MCU condition to initial
        #  to change the initial setting, just modify the items from here

        #  MCU will be in normal mode (EN, SW) = (1, 1) => 4
        self.pmic_mode(4)
        pass

    def pulse_out(self, pulse_1, pulse_2):
        self.pulse1 = pulse_1
        self.pulse2 = pulse_2
        self.mcu_write('swire')
        pass

    def pmic_mode(self, mode_index):
        self.mode_set = mode_index
        self.mcu_write('en_sw')
        pass

    def relay_ctrl(self, channel_index):
        self.meter_ch_ctrl = channel_index
        self.mcu_write('relay')
        pass

    def i2c_single_write(self, register_index, data_index):
        self.reg_i2c = register_index
        self.data_i2c = data_index
        self.mcu_write('i2c')

    # to update the implementation of other function
    # think about what is needed from the MCU operation
    # update later

    pass

#  testing items for the MCU control object


if __name__ == '__main__':

    pass
