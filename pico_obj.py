"""
this file is used to build up the Raspberry pi PICO ontrol object at
python side, which can be call by the other python program
PICO will be used to build up MCU related function
"""

import pyvisa

# # also for the jump out window, same group with win32con
import win32api
from win32con import MB_SYSTEMMODAL

# for float operation of string
import locale as lo

# for the delay function
import time

import serial

# turn off the formatter
# fmt: off

# pyvisa only support 9600 baud rate, use serial for higher baudrate
rm = pyvisa.ResourceManager()

class PICO_obj ():

    def __init__(self, sim_mcu0=0, com_addr0=0):
        # this is the initialize sub-program for the class and which will operate once class
        # has been defined

        self.sim_mcu = sim_mcu0
        self.com_addr = com_addr0
        if self.com_addr == 100:
            # set PICO to simulation mode if the addr is set to 100
            self.sim_mcu = 0

        # below meter mapping is for efficiency measurement
        # meter channel indicator: 0: Vin, 1: AVDD(BK), 2: OVDD(LDO), 3: OVSS(GND), 4: VOP(VCC), 5: VON(PG)
        # array mpaaing for the relay control
        self.meter_ch_ctrl = 0
        self.array_rst = '00'

        self.pulse1 = 0
        self.pulse2 = 0
        self.mode_set = 4
        # mode sequence: 1-4: (EN, SW) = (0, 0),  (0, 1), (1, 0), (1, 1) => default normal

        # i2c register and data (single byte data); slave address is fixed in the MCU
        # here only support for the register and data change
        self.reg_i2c = ''
        self.data_i2c = ''

        #  usb command string
        self.usb_cmd_str = ''


        pass

    def com_open(self):
        '''
        231112: no need to change the baud rate, default should be enough
        PICO use just like genetal communication method like input() of python
        '''
        # this function is used to open the com port of the MCU
        # this will be set independentlly in each object
        print('now the COM port of PICO is on')
        uart_cmd_str = f'COM{self.com_addr}'
        print(uart_cmd_str)
        if self.sim_mcu == 1:
            self.mcu_com = rm.open_resource(uart_cmd_str)
        else:
            print('open COM port but bypass the real operation')

        pass

    def com_query(self, cmd_str0=''):

        pass

    def write(self, command=0):
        self.mcu_com.query(command)

        pass

    def com_close(self):
        # after the verification is finished, reset all the condition
        # to initial and turn off the communication port
        self.back_to_initial()
        print('the MCU will turn off')
        if self.sim_mcu == 1:
            self.mcu_com.close()
        else:
            print('the com port is turn off now')

        pass

    def back_to_initial(self):
        # this sub program used to set all the MCU condition to initial
        # to change the initial setting, just modify the items from here

        # MCU will be in normal mode (EN, SW) = (1, 1) => 4
        self.pmic_mode(4)
        # the relay channel also reset to the default

        print('command accept to reset the MCU_PICO_grace')
        pass

    def pmic_mode(self, mode_index):
        '''
        (EN,SW) or (EN2, EN1) \n
        1:(0,0); 2:(0,1); 3:(1,0); 4:(1,1)
        '''
        # mode index should be in 1-4
        if mode_index < 1 or mode_index > 4:
            mode_index = 1
            # turn off if error occur
        self.mode_set = mode_index
        # not done yet.. decide after knowing what is the final decision of PICO side
        pass

if __name__ == '__main__':

    # testing for the PICO MCU

    # define the new PICO MCU and open the device

    sim_pcio = 1
    com_addr = 6
    baud_rate = 115200

    g_pico = PICO_obj(sim_mcu0=sim_pcio, com_addr0=com_addr)

    g_pico.com_open()

    testing_index = 1

    if testing_index == 0 :
        '''
        using testing MCU(TI) to send COM command and check from the LA bus
        '''
        ti_mcu_com = 6
        # virtual_com = rm.open_resource(resource_name=f'COM{ti_mcu_com}')

        # # pyvisa don't have change function, use COM default 9600
        # temp = virtual_com.baudrate
        # print(f'current baudrate: {temp}')
        # virtual_com.baudrate = 115200
        # temp = virtual_com.baudrate
        # print(f'update baudrate: {temp}')

        # another object is used to test the serial function
        virtual_com_s = serial.Serial(port=f'COM{ti_mcu_com}', baudrate=115200)
        temp = virtual_com_s.baudrate
        print(f'virtual_com_s: current baudrate: {temp}')
        virtual_com_s.baudrate = 9600
        temp = virtual_com_s.baudrate
        print(f'virtual_com_s: update baudrate: {temp}')
        # data_to_send = 'ABCDE12345'
        # encoded_data = data_to_send.encode('utf-8')
        # print(encoded_data)
        # virtual_com_s.write(encoded_data)

        while 1 :

            data_to_send = 'a'
            encoded_data = data_to_send.encode('utf-8')
            print(encoded_data)
            virtual_com_s.write(encoded_data)

            data_to_send = 'b'
            encoded_data = data_to_send.encode('utf-8')
            print(encoded_data)
            virtual_com_s.write(encoded_data)

            data_to_send = 'c'
            encoded_data = data_to_send.encode('utf-8')
            print(encoded_data)
            virtual_com_s.write(encoded_data)






        pass

    elif testing_index == 1 :

        while 1 :

            data_to_send = 'a'
            encoded_data = data_to_send.encode('utf-8')
            print(data_to_send)
            g_pico.write(data_to_send)

            data_to_send = 'b'
            encoded_data = data_to_send.encode('utf-8')
            print(data_to_send)
            g_pico.write(data_to_send)

            data_to_send = 'c'
            encoded_data = data_to_send.encode('utf-8')
            print(data_to_send)
            g_pico.write(data_to_send)


        pass
