import pyvisa
import time

rm = pyvisa.ResourceManager()

# turn off the formatter
# fmt: off

uart_cmd_str = 0
# UART command string
wait_time = 0.5
# unit in second
mcu_com_addr = 3
# MCU COM port number
mode = 1
# testing MCU mode settings
content_1 = 5
content_2 = 7

loop_out = 0
# loop control for COM port
temp_out = 0
# optaional part output

c_COM = 3
# the repeat of sending/ reading data

mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
# array for realy board

mcu_cmd_arry_2 = ['FF', '02', '04', '08', '10', '20', '40', '80']
# array for pattern gen, each bit maps to one output at P2OUT of MCU

mcu_cmd_arry_I2C = ['A0', 'F0']
i2c_com = 2

uart_cmd_str = "COM" + str(int(mcu_com_addr))
print(uart_cmd_str)
mcu_com = rm.open_resource(uart_cmd_str)

# test mode selecton
test_mode = 1
# general testing pattern is to send 3 pattern
# mode 1,2(SWIRE): mode + swire_pulse1 + swire_pulse2
# mode 3(IQ_measurement)  : mode + control_stste (from 1 - 4)
# mode 4(I2C mode): mode + HP_1 + LP_1 +HP_2 + LP_2
# mode 5(RA mode): mode + P2OUT(relay)
# mode 6(pattern generator): mode + TBD (this should based on the implementation method of pattern gen in summary file)
# test mode build in another file (20211202)

# watch out the form of the address using for connect COM port or GPIB

if temp_out == 0:
    # testing for the python dynamic pattern gen
    if loop_out == 1:
        x_COM = 0
        while x_COM < c_COM:

            # MCU initial for IQ test under normal mode
            uart_cmd_str = chr(mode) + chr(content_1) + chr(content_2)
            print(uart_cmd_str)
            mcu_com.write(uart_cmd_str)
            time.sleep(wait_time)
            # input()
            print('loop counter now is:' + str(x_COM))

            # write example
            # mcu_com.write('abcde')
            # read example
            # mcu_res = mcu_com.read()
            x_COM = x_COM + 1

    else:

        uart_cmd_str = chr(mode) + chr(content_1) + chr(content_2)
        print(uart_cmd_str)
        mcu_com.write(uart_cmd_str)
        time.sleep(wait_time)
        # input()


elif(temp_out == 1):
    # testing for the python dynamic pattern gen (pattern2)
    # temp output

    # x_COM = 0
    # while x_COM < c_COM:

    #     # MCU initial for IQ test under normal mode
    #     uart_cmd_str = chr(mode) + mcu_cmd_arry[x_COM]
    #     print(uart_cmd_str)
    #     mcu_com.write(uart_cmd_str)
    #     time.sleep(wait_time)
    #     # input()
    #     print('loop counter now is:' + str(x_COM))

    #     # write example
    #     # mcu_com.write('abcde')
    #     # read example
    #     # mcu_res = mcu_com.read()
    #     x_COM = x_COM + 1

    # contimuous sending command to see the minimum time needed for this mode

    x_COM = 0
    while x_COM < c_COM:

        # MCU initial for IQ test under normal mode
        uart_cmd_str = chr(mode) + mcu_cmd_arry[x_COM]

        # print(uart_cmd_str)
        mcu_com.write(uart_cmd_str)

        uart_cmd_str = chr(mode) + mcu_cmd_arry[x_COM+1]

        mcu_com.write(uart_cmd_str)
        # time.sleep(wait_time)
        # input()
        # print('loop counter now is:' + str(x_COM))

        # write example
        # mcu_com.write('abcde')
        # read example
        # mcu_res = mcu_com.read()
        x_COM = x_COM + 1

elif(temp_out == 2):
    # testing for the I2C function in MCU python version
    x_COM = 0
    uart_cmd_str = ''
    while x_COM < i2c_com:

        # MCU initial for IQ test under normal mode
        uart_cmd_str = uart_cmd_str + mcu_cmd_arry_I2C[x_COM]
        x_COM = x_COM + 1
        # print(uart_cmd_str)
    uart_cmd_str = chr(4) + uart_cmd_str
    print(uart_cmd_str)
    # i2C is at mode 4
    mcu_com.write(uart_cmd_str)

    # time.sleep(wait_time)
    # input()
    # print('loop counter now is:' + str(x_COM))

    # write example
    # mcu_com.write('abcde')
    # read example
    # mcu_res = mcu_com.read()
