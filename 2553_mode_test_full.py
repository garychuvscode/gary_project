import pyvisa
import time

rm = pyvisa.ResourceManager()

uart_cmd_str = 0
# UART command string
wait_time = 0.5
# unit in second
mcu_com_addr = 13
# mcu_com_addr = 3
# MCU COM port number
mode = 0
# testing MCU mode settings

c_COM = 8
# the repeat of sending/ reading data

mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
# array for relay board

mcu_cmd_arry_2 = ['FF', '02', '04', '08', '10', '20', '40', '80']
# array for pattern gen, each bit maps to one output at P2OUT of MCU

mcu_cmd_arry_I2C_org = ['C3', '1C', 'C5', '05', 'C6', '1B', 'C4', '2B']
mcu_cmd_arry_I2C_0p1 = ['C3', '1D', 'C5', '06', 'C6', '1C', 'C4', '2C']
# array element: register + data as one group
i2c_single_byte = 2
# one data with one command
c_i2c_byte = 4
# 374 as example using 4 byte of change then output change
select_i2c = 1
# 0 is the org array and 1 is 0p1 array

# SWIRE testing (two pulse used for the testing)
pulse1 = 5
pulse2 = 15

# start up for the COM port from pyvisa
uart_cmd_str = "COM" + str(int(mcu_com_addr))
print(uart_cmd_str)
mcu_com = rm.open_resource(uart_cmd_str)

# general support code

#     # write example
#     # mcu_com.write('abcde')
#     # read example
#     # mcu_res = mcu_com.read()
#     x_COM = x_COM + 1

# end of support code

# test mode selecton
test_mode = 5
test_mode_pat = 0
# 0 include the start command, 1 only the stop command
# general testing pattern is to send 3 pattern
# mode 1,2(SWIRE): mode + swire_pulse1 + swire_pulse2
# mode 3(IQ_measurement)  : mode + control_stste (from 1 - 4)
# mode 4(I2C mode): mode + HP_1 + LP_1 +HP_2 + LP_2
# mode 5(RA mode): mode + P2OUT(relay)
# mode 6(pattern generator V1-5ms): mode + P2OUT(function gen, no relay delay)
# mode 7(pattern generator V2-100us): mode + pattern selection + don't care
# mode 8(pattern generator V3-5us): mode + pattern selection +minimum clock selection
# implementation method of pattern gen in summary file)
# test mode build in another file (20211202)


# watch out the form of the address using for connect COM port or GPIB

if test_mode == 1:
    # here is for SWIRE testing
    # SWIRE initialization
    mode = 1

    # MCU initial for IQ test under normal mode
    uart_cmd_str = chr(mode) + chr(pulse1) + chr(pulse2)
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)
    time.sleep(wait_time)
    print('send finished with pulse ' + str(pulse1) + 'and ' + str(pulse2))

elif test_mode == 2:
    # here is for SWIRE testing
    # SWIRE initialization
    mode = 2

    # MCU initial for IQ test under normal mode
    uart_cmd_str = chr(mode) + chr(pulse1) + chr(pulse2)
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)
    time.sleep(wait_time)
    print('send finished with pulse ' + str(pulse1) + ' and ' + str(pulse2))

elif test_mode == 3:
    # here is for IQ testing
    mode = 3
    # mode sequence: 0-3: (EN, SW) = (0, 0),  (0, 1), (1, 0), (1, 1)
    x_submode = 0
    # initialize the counter for differnt state
    while x_submode < 4:
        # submode + 1 will be the state command for MCU
        uart_cmd_str = (chr(int(mode)) +
                        chr(x_submode + 1) + chr(1))
        # for IQ scan only need one state command to MCU
        print(uart_cmd_str)
        mcu_com.write(uart_cmd_str)
        time.sleep(wait_time)
        if x_submode == 0:
            time.sleep(wait_time)
            # extra wait time for the normal mode to shtudown mode transition
            # because IQ may not stop change so fast, need to double check
        x_submode = x_submode + 1
        input()

        # may need to add the I/O reset pin for EN and SWIRE after changing the mode
        # for either efficiency measurement or the testiing using mode 3!

elif test_mode == 4:
    # here is for I2C testing
    mode = 4

    x_i2c = 0
    # initialize the counter for differnt state
    while x_i2c < 4:
        if select_i2c == 1:
            uart_cmd_str = (chr(int(
                mode)) + mcu_cmd_arry_I2C_0p1[2 * x_i2c] + mcu_cmd_arry_I2C_0p1[2 * x_i2c + 1])
        else:
            # choose different array for the I2C command output
            uart_cmd_str = (chr(int(
                mode)) + mcu_cmd_arry_I2C_org[2 * x_i2c] + mcu_cmd_arry_I2C_org[2 * x_i2c + 1])
        print(uart_cmd_str + str(select_i2c))
        mcu_com.write(uart_cmd_str)
        time.sleep(wait_time)
        input()

        x_i2c = x_i2c + 1

elif test_mode == 5:
    # here is for the RA mode testing (relay board control)
    mode = 5

    c_item = 0
    relay_amount = 8
    while c_item < relay_amount:

        uart_cmd_str = chr(int(mode)) + mcu_cmd_arry[c_item]
        mcu_com.write(uart_cmd_str)
        print(uart_cmd_str)
        time.sleep(wait_time)
        input()

        c_item = c_item + 1

    # the reset to initial state
    print('going to back to initial state')
    input()
    uart_cmd_str = chr(int(mode)) + '00'
    mcu_com.write(uart_cmd_str)
    # watch out that the reset responsibility will be at python(PC) side
    # since the resource is less limitation here

elif test_mode == 6:
    # this is for pattern generator V1

    mode = 6

    c_cycle = 0
    pattern_amount = 8
    cyc_amount = 4
    while c_cycle < cyc_amount:
        c_pattern = 0
        while c_pattern < pattern_amount:

            uart_cmd_str = chr(int(mode)) + mcu_cmd_arry_2[c_pattern]
            mcu_com.write(uart_cmd_str)
            print(uart_cmd_str)
            # to see the change of pattern in minimum clock time, put two pattern change at the same time
            # to prevent overflow of the array, add the command for second pattern limit when c_pattern
            # is smaller than 6, but last pattern don't have pattern change to see the change
            if c_pattern < 6:
                uart_cmd_str = chr(int(mode)) + mcu_cmd_arry_2[c_pattern+2]
                mcu_com.write(uart_cmd_str)
            print(uart_cmd_str)
            time.sleep(wait_time)
            # input()

            c_pattern = c_pattern + 1

        c_cycle = c_cycle + 1
        # input()

elif test_mode == 7:
    # this is for pattern generator V2

    mode = 7
    if test_mode_pat == 0:
        # first to send the trigger to start output of the pattern gen
        uart_cmd_str = chr(int(mode)) + chr(int(3)) + chr(int(1))
        print(uart_cmd_str)
        mcu_com.write(uart_cmd_str)

        #  after enable the pattern gen, send the disable command
        #  see if the output can stop after control is send
        input()
    uart_cmd_str = chr(int(mode)) + chr(int(4)) + chr(int(1))
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)

elif test_mode == 8:
    # this is for pattern generator V3, same with V2, only mode change

    mode = 8

    if test_mode_pat == 0:
        # first to send the trigger to start output of the pattern gen
        uart_cmd_str = chr(int(mode)) + chr(int(1)) + chr(int(1))
        print(uart_cmd_str)
        mcu_com.write(uart_cmd_str)

        #  after enable the pattern gen, send the disable command
        #  see if the output can stop after control is send
        input()
    uart_cmd_str = chr(int(mode)) + chr(int(4)) + chr(int(1))
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)
