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

        # 230920 add I2C bit write function
        # 231128 copy from JIGM3 for the configuration of bit process

        self.bit_clr = {
            "0": 0xFE,
            "1": 0xFD,
            "2": 0xFB,
            "3": 0xF7,
            "4": 0xEF,
            "5": 0xDF,
            "6": 0xBF,
            "7": 0x7F,
        }

        self.bit_set = {
            "0": 0x1,
            "1": 0x2,
            "2": 0x4,
            "3": 0x8,
            "4": 0x10,
            "5": 0x20,
            "6": 0x40,
            "7": 0x80,
        }


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
            check_ID = ''
            while 1 :
                # 240117 add new operation: clear the resource for reset
                list_dev = rm.list_resources()
                print(f'the related resource we have \n {list_dev}')
                try:
                    # first use the default value
                    self.mcu_com = rm.open_resource(uart_cmd_str,baud_rate=baud_rate0)
                    check_ID = self.p_query(cmd_str0="*IDN?")
                    print(f'MCU ID_check finished, pico and {check_ID}')
                    # 240128: use try to check if the resource open ok

                    # for the connection speed, at least use 115200, there
                    # are time out issue if the frequency is too slow
                    if check_ID == 'Grace':
                        print(f'correct and break')
                        break
                    else:
                        print(f'ID not correct')
                        raise Exception("ID not correct")  # 使用 raise 引发异常

                except Exception as e:
                    # search from all USB device
                    available_devices = rm.list_resources()

                    for device in available_devices:
                        print(f'now test: {device} ')
                        try:
                            # to check what is pico
                            self.mcu_com = rm.open_resource(str(device))
                            self.mcu_com.clear()
                            self.mcu_com.write("*IDN?")
                            # the first of read in pico after write is to get command
                            cmd_write = self.mcu_com.read()
                            # the second read is the return item (if there are return)
                            item_back = self.mcu_com.read()
                            check_ID = item_back.strip()
                            print(item_back)
                            print(
                                f"what we got on usb is: first the command {cmd_write},second the item_back {item_back}"
                            )
                            if check_ID == 'Grace':
                                print(f'correct and break')
                                # this break if active for the for loop

                                break

                        except Exception as e:
                            # may not have or wrong device
                            print(f"exception: {e}, please check pico connection")
                            pass



                        # end of check device for loop
                        pass

                    if check_ID == 'Grace':
                        print(f'correct and break')
                        break
                    # end of the external exception
                    pass

                if check_ID != 'Grace' :
                    self.message_box(content_str=f'the ID: "{check_ID}" is wrong, check PCIO MCU connection', title_str='MCU not found')
                # end of while
                pass



        else:
            print('open COM port but bypass the real operation')

        pass

    def message_box(self, content_str, title_str, box_type=0):
        '''
        message box function
        auto_exception is for waveform capture, will bypass fully auto setting in global setting \n
        boxtype(mpaaed with return value): 0-only confirm\n
        1-confirm: 1, cancel: 2
        2-stop: 3, re-try: 4, skip: 5
        3-yes: 6, no: 7, cancel: 2
        4-yes: 6, no: 7
        '''
        content_str = str(content_str)
        title_str = str(title_str)
        msg_res = 7
        # won't skip if not enter the result update
        msg_res = win32api.MessageBox(0, content_str, title_str, box_type)
        # 0 to 3 is different type of message box and can sen different return value
        # detail check on the internet
        print('msg box call~~ ')
        print('P.S Grace is cute! ~ ')

        return msg_res

    def p_query(self, cmd_str0='t;usb;2', time_out_s0 =5):
        '''
        command send and looking for feedback => used for
        data return of checking the feedback of command
        send to PICO(double check)
        test_LED 2 => GP9
        240124: better to clear the COM port RX FIFO before query
        to get the latet information => using read or clear
        '''

        if self.sim_mcu == 1 :
            try:
                # try to read out all the stuff first
                tmp_r = self.mcu_com.read()
                print(f'first to clear FIFO: {tmp_r}')

            except Exception as e :
                print(f'FIFO is empty or other error: {e}')

            try:
                # then start query
                tmp_cmd = self.mcu_com.query(cmd_str0, time_out_s0)
                # 240128 add one more offset for the read offset
                tmp_r = self.mcu_com.read()
                tmp_r = tmp_r.strip()
                print(f'pico repeat item send: {tmp_cmd}, result back: {tmp_r}')
                return tmp_r
                pass
            except Exception as e :
                print(f'query error, need to check command with error: {e}')
                pass

            pass
        else:
            print(f'PICO sim_mode query {cmd_str0}')

        # end of query
        pass



    def p_write(self, command='t;usb;5'):
        '''
        run the command and without return
        only try to print the result
        test_LED 5 => GP22
        240124 watch out that the old commend send to PICO will be
        save in the FIFO buffer if sending command faster than
        excuting
        '''
        ret_from_pico = 0
        try:
            if self.sim_mcu == 1:
                ret_from_pico = self.mcu_com.write(command)
                # ret_from_pico = self.mcu_com.query(command)
            else:
                ret_from_pico = f'sim_mode_{command}'
            print(f'Grace is about 30y, and she say: {ret_from_pico}')
        except Exception as e :
            print(f'write error "{e}" at COM{com_addr}')

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
        self.p_write('grace;io_rst')

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

    def i2c_read(self):
        '''
        define for I2C read, need more modify
        '''

        # int_list = struct.unpack("<BBBB", byte_data)
        # 231128 this is used for i2c read byte to number transfer
        # refer to byte_list to int_list in chat GTP
        pass

    def i2c_write(self, num0=0, status0=0):

        pass

    def relay_ctrl(self, channel_index=0, relay_mode0=1, t_dly_s=0):

        pass

    def pico_gio(self, num0=0, status0=0):
        '''
        240125 the IO function now is different from different MCU
        But it should be ok if going to have same function call by main
        it can be implement in another function, this should be ok
        now just use the own function
        '''



        pass

    def pulse_out(self, pulse_1=1, pulse_2=1):
        """
        PICO mapped with MSP430
        pulse need to be less then 255
        can choose pulse at GP6(EN), default is at GP7(SW)

        PICO duration is 10 us fixed
        """
        """
        first to use pulse gen, for the PIO, will be in
        future plan
        """

        pass

    def hex_to_num(self, data_in0=0):
        '''
        try to transfer the input to number for register write
        skip if already number (but watch out this is dec number input, not hex)
        '''
        try:
            data_in0 = int(data_in0, 16)
            print(f'input is string and transfer number: {data_in0}')
        except:
            print(f'input is already number: {data_in0}, return origin')

        return data_in0

    def num_to_hex (self, data_in0=0):
        '''
        transfer the number to hex
        '''
        try:
            data_in0 = hex(data_in0)
            return data_in0
        except:
            print(f'input data "{data_in0}" transfer error, double check input data')
            return 'data_error'

        pass

    def bit_s(self, bit_num0=0, byte_state_tmp0=0):
        bit_num0 = str(bit_num0)

        # bit set process
        bit_cmd0 = self.bit_set[bit_num0]
        print(f"set bit_cmd0 is {bit_cmd0}")
        new_byte_data = byte_state_tmp0 | bit_cmd0
        print(f"final command0 is {new_byte_data}, g")

        return new_byte_data

    def bit_c(self, bit_num0=0, byte_state_tmp0=0):
        bit_num0 = str(bit_num0)

        # bit clear process
        bit_cmd0 = self.bit_clr[bit_num0]
        print(f"clear bit_cmd0 is {bit_cmd0}")
        new_byte_data = byte_state_tmp0 & bit_cmd0
        print(f"final command0 is {new_byte_data}, g")

        return new_byte_data

    def pure_group_write(self, lsb0=0, len0=1, data0=0, byte_state_tmp0=0):

        '''
        group few bits together for register adjustment in program
        lsb is define the LSB of this group (from 0-7)
        len is length of this group (from 1-8)
        limitation is all the bit must be in same register
        data need to be integer

        byte_state_tmp0 is original data, and data0 is the new data
        '''
        lsb0 = int(lsb0)
        len0 = int(len0)
        data0 = int(data0)
        byte_state_tmp0 = self.hex_to_num(byte_state_tmp0)
        '''
        calculation method register update:
        left shift for LSB: LSB0 => no need shift, LSB3 => give three 0 in right
        '''
        # x**y operator is means x^y, sinc the ^ means XOR in python
        if data0 > 2**(len0) :
            # data is too big, output the error message and bypass the command
            # this check is used to prevent overflow of Grace XD
            print(f'length "{len0}" and data "{data0}" have fconflict, please double check ')

            pass
        else:
            # shift the data based on the length and using XOR to update result
            # then write the new data to the register
            group_command = data0 << lsb0
            print(f'Gary want new data to be: {hex(group_command)}, {bin(group_command)}')

            temp = 0
            for x in range(len0):
                temp = temp + 2**x
            temp = temp << lsb0
            print(f' what Grace said: {bin(temp)} ~')
            temp = 255 - temp
            print(f' what Grace do: {bin(temp)} XDD')

            # byte_state_tmp = self.i2c_read(device=device0, regaddr=register0, len=1)
            # byte_state_tmp = byte_state_tmp[0]
            # old one is used for I2C mode, update new method, the old datais
            # this data is from function input
            byte_state_tmp = byte_state_tmp0

            print(f'Grace get original data: {hex(byte_state_tmp)}, {bin(byte_state_tmp)}')
            # clean the group bits before adding the new data
            byte_state_tmp = byte_state_tmp & temp
            print(f'Gary clean original data: {hex(byte_state_tmp)}, {bin(byte_state_tmp)}')
            # using or is ok after clean the related group bits
            new_byte_data = group_command | byte_state_tmp
            x = hex(new_byte_data)
            print(f"Grace's final command0 is {x}")


            # 231018: take off the i2C write from the i2c, this function is only for byte operation

            # # write the new byte to the register, change to JIGM3 format first
            # new_list = [new_byte_data]

            # self.i2c_write(device=device0, regaddr=register0, datas=new_list)

            # # double check process
            # check_data = self.i2c_read(device=device0, regaddr=register0, len=1)
            # check_data = check_data[0]
            # print(f'Grace read back from reg {hex(register0)} have new data {hex(check_data)}')

            # return the calculation result
            return new_byte_data


        pass

if __name__ == '__main__':

    # testing for the PICO MCU

    # define the new PICO MCU and open the device

    sim_pcio = 1
    # 230117: for the COM port selection => based on the
    # hardware manager, or the NI_MAX
    com_addr = 1
    baud_rate0 = 115200

    g_pico = PICO_obj(sim_mcu0=sim_pcio, com_addr0=com_addr)

    g_pico.com_open()

    testing_index = 2

    if testing_index == 0 :
        '''
        using testing MCU(TI) to send COM command and check from the LA bus
        240124 add comments, this is just testing the com port communication
        USB can be use directly the communication
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
        '''
        this is also old code may not be use
        just leave for record
        '''
        while 1 :

            data_to_send = 'a'
            encoded_data = data_to_send.encode('utf-8')
            print(data_to_send)
            g_pico.p_write(data_to_send)

            data_to_send = 'b'
            encoded_data = data_to_send.encode('utf-8')
            print(data_to_send)
            g_pico.p_write(data_to_send)

            data_to_send = 'c'
            encoded_data = data_to_send.encode('utf-8')
            print(data_to_send)
            g_pico.p_write(data_to_send)


        pass

    elif testing_index == 2 :
        '''
        240117 checking the connection of USB to PICO
        through USB directly, and blink LED for command check
        '''
        x = 0
        while x < 5 :
            g_pico.p_write()
            time.sleep(3)
            x = x + 1
        pass
