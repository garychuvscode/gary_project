# this file is mainly for the testing of coding


import locale as lo
import sys
import time
from datetime import datetime

#  add MCU testing for old version board
import mcu_obj as m

import win32api
from win32con import MB_SYSTEMMODAL

import xlwings as xw

mcu_s = m.MCU_control(0, 3)
mcu_s.com_open()

# turn off the formatter
# fmt: off

class test_calass():

    def __init__(self):

        pass

    def scope_ch(self):

        self.ch_c1 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': 20, 'filter': 2, 'v_offset': 0,
                      'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}
        self.ch_c2 = {'ch_view': 'TRUE', 'volt_dev': '0.5', 'BW': 20, 'filter': 2, 'v_offset': 0,
                      'label_name': 'name', 'label_position': 0, 'label_view': 'TRUE', 'coupling': 'DC1M'}

        for i in range(1, 8+1):
            if i == 1:
                temp_dict = self.ch_c1

            a = temp_dict['ch_view']
            print(f'app.Acquisition.C{i}.View = {a}')
            print(f'app.Acquisition.C{i}.View = {temp_dict["ch_view"]}')

        pass

    def two_dim_dict(self):

        self.p1 = {"param": "max", "source": "C1", "view": "TRUE"}
        self.p2 = {"param": "min", "source": "C2", "view": "FALSE"}

        self.mea_ch = {"P1": self.p1, "P2": self.p2, }

        # multi layer of the list (dictionary)

        temp = list(self.mea_ch)[0]
        # 'P1'
        temp3 = list(self.mea_ch.values())[0]
        # self.p1
        trmp2 = list(list(self.mea_ch.values())[0])[0]
        # 'param'
        temp4 = (list(self.mea_ch.values())[0])["param"]
        # 'max'
        list(self.mea_ch.values())[0] = self.p2
        # this line can't be done
        self.mea_ch['P1'] = self.p2
        # this is the method to assign

        print(
            f"app.Measure.{list(self.mea_ch)[0]}.ParamEngine = '{(list(self.mea_ch.values())[0])['param']}'")

        pass

    def float_format(self):

        # this fuction used to test format output for the float
        a = 0.123456789
        dig = 5
        b = float("{:.2f}".format(a))
        # c = float(f"{:.{dig}f}".format(a)) => error format
        d = "{:.4f}".format(a)
        e = lo.atof("{:.2f}".format(a))

        print(b)

        pass

    def float_gene(self, input, scaling=1000, digit=2):
        '''
        transfer the digit of float \n
        input can be string or float \n
        default scaling to mV
        '''
        a = float(input)
        a = a * scaling
        if digit == 0:
            b = float("{:.0f}".format(a))
        elif digit == 1:
            b = float("{:.1f}".format(a))
        elif digit == 2:
            b = float("{:.2f}".format(a))
        elif digit == 3:
            b = float("{:.3f}".format(a))
        elif digit == 4:
            b = float("{:.4f}".format(a))
        print(b)
        return b

    def time_string(self):
        # refer to the command from internet to check
        now = datetime.now()  # current date and time

        year = now.strftime("%Y")
        print("year:", year)

        month = now.strftime("%m")
        print("month:", month)

        day = now.strftime("%d")
        print("day:", day)

        time = now.strftime("%H:%M:%S")
        print("time:", time)

        date_time = now.strftime("%Y-%m-%d, %H:%M:%S")
        print("date and time:", date_time)

        date_stamp = now.strftime("%Y_%m_%d_%H_%M")
        print('the stanmp: ' + date_stamp)

        pass

    def message_box_testing(self):

        response = win32api.MessageBox(
            0, "Did you hear the Buzzer?", "Buzzer Test", 4, MB_SYSTEMMODAL)
        # answer yes: 6, no: 7

        return response

    def message_box_testing(self):

        response = win32api.MessageBox(
            0, "Did you hear the Buzzer?", "Buzzer Test", 1)
        # answer confirm: 1, cancel: 2

        return response

    def message_box_testing(self):

        response = win32api.MessageBox(
            0, "Did you hear the Buzzer?", "Buzzer Test", 2)
        # answer stop: 3, re-try: 4, skip: 5

        return response

    def message_box_testing(self):

        response = win32api.MessageBox(
            0, "Did you hear the Buzzer?", "Buzzer Test", 3)
        # answer yes: 6, no: 7, cancel: 2

        return response

    def message_box_testing(self):

        response = win32api.MessageBox(
            0, "Did you hear the Buzzer?", "Buzzer Test", 0)
        # answer confirm: 1

        return response

    def message_box_testing(self):

        response = win32api.MessageBox(
            0, "Did you hear the Buzzer?", "Buzzer Test", 5)
        # answer re-try: 4, cancel: 2

        return response

    def try_except(self, example=0):

        # practice for try except application
        try:
            a = 0
            print(f'correct input answer is {a}')
            a = 100 / example
            print(f'2_correct input answer is {a}')
            pass
        except:
            a = 10
            print(f'this is wrong, enter error handling a become {a}')

            pass

    def logic_operator(self):
        '''
        try all the logic operator and result show in terminal
        '''

        x = 0x01
        y = 0x20
        print(x + y)

        print(x & y)
        print(x | y)
        print(x ^ y)
        print(~x)
        print(~y)

        print(hex(x))
        print(hex(y))

        print(hex(~x))
        print(hex(~y))

        print(hex(abs(~x)))
        print(hex(abs(~y)))

        a = 0b1010111111111111
        print(hex(a))

        pass

    def copy_range_test(self):
        # assign the destination range and add into the source range.copy
        # function to copy related ranage to the dest range

        '''
        for the input from the excel, using format in L1 ..
        for the input from program, using the index of row and column
        this should be the better way of implementation
        '''

        temp_range = self.sh_dest.range((3, 10), (10, 10))
        self.sh_sor.range((3, 10), (10, 10)).copy(temp_range)

        ind_start = 'L3'
        ind_end = 'L10'

        self.sh_sor.range(ind_start, ind_end).copy(temp_range)
        time.sleep(0.2)

        ind_start = (3, 10)
        ind_end = (10, 10)

        self.sh_sor.range(ind_start, ind_end).copy(temp_range)
        time.sleep(0.2)

        ind_start = 'L3'
        ind_end = 'L10'

        self.sh_sor.range(ind_start, ind_end).copy(temp_range)
        time.sleep(0.2)

        a = 3
        b = 10

        ind_start = (b, a)
        ind_end = self.cell_gen(a, b)

        self.sh_sor.range(ind_start, ind_end).copy(temp_range)

        pass

    def cell_gen (self, a0=0, b0=0):
        '''
        use to gen the range setting fom row, column input
        and return the list(array) of format needed to
        operate range
        '''
        print(f'row setting is {a0}')
        print(f'column setting is {b0}')


        return (a0, b0)

    def excel_open(self):
        '''
        open the excel and copy related range to destination
        '''
        self.wb = xw.Book()
        self.wb2 = xw.Book()
        self.sh_sor = self.wb.sheets.add('test_source')
        self.sh_dest = self.wb2.sheets.add('test_dest')
        self.sh_sor.range((10, 10), (3, 10)).value = '1010abc'

        self.sh_sor.range('L3', 'L10').value = '09876'

        temp_range = self.sh_sor.range('L3', 'L10')

        pass

    def hex_check(self):

        '''
        239018 check for hex number input from excel and output to the JIGM3
        '''

        a = 231
        b = 32
        c = 'FF'
        d = 'ff'

        # transfer the string to hex number
        '''
        all the number is just int, the difference is only how to display
        => so just transfer to the integer
        '''
        x = int(c,16)
        x1 = int(d,16)

        print(x)
        print(x1)

        print(hex(x))
        print(hex(x1))

        # format without 0x, this is the string
        y=hex(x)[2:]
        y1=hex(x1)[2:]

        print(y)
        print(y1)

        '''
        for MSP430, just take the data from excel, and output to
        the I2C write and read

        for JIGM3, get the string and transfer to the integer by using
        x1 = int(d,16) and save into the list
        '''

        # transfer the data to list of reading


        pass




t_s = test_calass()
testing_index = 9

if testing_index == 0:
    print('a')
    # from 0-9 => < 10 and start from 0 (like array)
    for i in range(10):
        print(i, end=' ')
        # not change line

    for i in range(1, 1+8):
        print(i)

        t_s.scope_ch()

elif testing_index == 1:
    # testing for the 2 dimension dictionary

    t_s.two_dim_dict()
    t_s.float_format()
    t_s.float_gene('0.123456789', 100)
    t_s.float_gene('0.123456789', digit=0)
    t_s.float_gene('0.123456789', digit=1)
    t_s.float_gene('0.123456789', digit=2)
    t_s.float_gene('0.123456789', digit=3)
    t_s.float_gene('0.123456789', digit=4)

    pass

elif testing_index == 2:

    # testing for exit function of system

    x = 0
    while x < 100:

        if x == 50:
            sys.exit()

        print(f'now is {x}')

        x = x + 1

        pass

    pass

elif testing_index == 3:

    # testing for exit function of system

    x = 0
    while x < 80000:
        a = 0
        while a < 8:
            mcu_s.relay_ctrl(a)
            # testing for the crash of relay control @ MCU
            print(a)
            a = a+1

        x = x+1
        print(f'x now is {x}')
        time.sleep(0.3)
        pass

    pass

elif testing_index == 4:

    t_s.time_string()

    pass

elif testing_index == 5:
    # testing for the message box and use the return to decide to skip the line or not

    for i in range(3):
        a = t_s.message_box_testing()
        print(a)

    pass

elif testing_index == 6:

    t_s.try_except(1)
    t_s.try_except(2)
    t_s.try_except(3)
    t_s.try_except(4)

    t_s.try_except(0)
    t_s.try_except(1)

    pass

elif testing_index == 7:

    t_s.logic_operator()

    pass

elif testing_index == 8:

    t_s.excel_open()
    input()
    t_s.copy_range_test()

    pass

elif testing_index == 9:

    t_s.hex_check()

    pass
