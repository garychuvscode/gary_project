# import for the object and package needed

# mode control variable, simulation or real(with instrument)

# from win32con import MB_SYSTEMMODAL, PAN_WEIGHT_BLACK

from win32con import MB_SYSTEMMODAL
import win32api
import time
import pyvisa
import xlwings as xw

import locale as lo
# include for atof function => transfer string to float

sim_real = 1
# 1 => real mode, 0=> simulation(program debug) mode
sim_mcu = 1
# this is only used for MCU com port test simuation
# when setting (sim_real, sim_mcu) = (0, 1)
# enable MCU testing when GPIC disable
# because NCU will be separate with GPIB for implementation and test

mcu_cmd_arry = ['02', '04', '08', '10', '20', '01', '40', '80']
mode = 5
# MCU mapping for RA_GPIO control is mode5


control_book_trace = 'c:\\py_gary\\test_excel\\RA_control.xlsx'
result_book_trace = 'c:\\py_gary\\test_excel\\RA.xlsx'

if sim_real == 1:
    import package_test_real

# for excel related function
# for UART and GPIB related function
# for delay related function
# for message box and others
# other function for message

# the way to import object seems to be pretty easy,
# but need to be in same folder with original python code

# this will cause error at the initialization of simulation mode
# if sim_real == 1 :
#     from package_test_real import LPS_505N
#     from package_test_real import Met_34460


# can think about the way of import:
# maybe not only the object(class) but also function
# and variable can be import from the other file in the same folder?


# talking block ~~ below for reference
# response = win32api.MessageBox(0, "Did you hear the Buzzer?", "Buzzer Test", 2, MB_SYSTEMMODAL)
# response2 = win32api.MessageBox(0, 'hello', 'title')

# ======== GPIB and UART initialization
# consider to put the control of instrument into sub program in other file

rm = pyvisa.ResourceManager()
# load the GPIB number and UART com port number like previous excel
# and IDN can be used to get related device name
# loading the GPIB address for instrument

# UART give the RA pulse to setting voltage
# UART also need to control the relay output
# UART control is equal to MCU control (GPIO for relay board)

# GPIB control the supply setting and measurement
# only 1 meter is needed, Iin can be done by power supply
# GPIB part power supply
# Vin open loop setting
# Iin measurement
# GPIB part meter
# voltage measurement, need delay and channel selection


# ======== GPIB and UART initialization end

# plan for instrument object definition

# ======== excel book and sheet operation
# control book loading => new book creation => default parameter loading

# open control workbook
# wb = xw.Book()

wb = xw.Book(control_book_trace)
# wb = xw.Book('E:\\RA_control.xlsx')

# wb = xw.Book('RA_control.xlsx')
# define new result workbook
wb_res = xw.Book()
# create reference sheet (for sheet position)
sh_ref = wb_res.sheets.add('ref_sh')
# delete the extra sheet from new workbook, difference from version
wb_res.sheets('工作表1').delete()

# define the sheets in control book
sh_main = wb.sheets('main')
sh_org_tab = wb.sheets('RA_output')
# copy the sheets to new book
sh_main.copy(sh_ref)
sh_org_tab.copy(sh_ref)

# assign both sheet to the new sheets in result book
sh_main = wb_res.sheets('main')
sh_org_tab = wb_res.sheets('RA_output')
# sh_main = sh_main.copy(sh_ref)
# sh_org_tab = sh_org_tab.copy(sh_ref)
# delete reference after copied
sh_ref.delete()

# save the result book and turn off the control book
wb_res.save(result_book_trace)
wb.close()
# close control books

# base on output format copied from the control book
# start parameter initialization
total_board = sh_main.range('G3').value
pulse_ELVDD = sh_main.range('G4').value
pulse_ELVSS = sh_main.range('G5').value
pulse_AVDD = sh_main.range('G6').value
before_after = sh_main.range('G7').value
dynamic_NO = sh_main.range('G8').value
Vin = sh_main.range('G9').value
dynamic_board = sh_main.range('G10').value
pwr_max_I = sh_main.range('G15').value
pwr_use_ch = sh_main.range('G16').value
pre_vin = sh_main.range('G17').value

# COM port parameter input
mcu_com_addr = sh_main.range('G14').value


# GPIB instrument related pa
# rameter loading
PWR_sup_addr = sh_main.range('G30').value
met_vout_addr = sh_main.range('G31').value

# variable definition explanation
# c_ => usually used for counter for loop

# measured vin simulation variable
v_sim = 0
# items counter (from RA_output, total 6, 1 current + 5 voltage)
c_item = 0
RA_out_V = 6
# board counter
c_board = 0
# x_start used to separate before and after at measurement
x_start = 0
# y_start used to separate before and after at measurement
y_start = 0
# read from excel if board number change (2 is used to save the old one)
c_dy_board = 0
c_dy_board2 = 0
first_check = 0
c_wait_time = 0
# python sleep time (unit in second)
wait_time = 0.5

# message box return result (most time is int), just for testing fthe msg box result
msg_res = 0

# ======== excel book and sheet operation


# GPIB an COM port initialization (hardware device)
if sim_real == 1:
    # initialization for the instrument used ( with reference address from excel )
    # if the import is the form of " from package_test_real import LPS_505N " or
    # " from package_test_real import Met_34460 "
    # PWR_supply1 = LPS_505N(0, 0, 1, PWR_sup_addr, 'off')
    # M1_v_in = Met_34460(0.0001, 7, 0.000001, 1, met_vout_addr)

    PWR_supply1 = package_test_real.LPS_505N(0, 0, 1, PWR_sup_addr, 'off')
    M1_v_in = package_test_real.Met_34460(
        0.0001, 15, 0.000001, 1, met_vout_addr)
    # open the GPIB port with related address from excel sheet

    PWR_supply1.open_inst()
    M1_v_in.open_instr()
    # update the instrument name to the main sheet
    sh_main.range((30, 8)).value = PWR_supply1.inst_name()
    sh_main.range((31, 8)).value = M1_v_in.inst_name()

    # COM port initial settings
uart_cmd_str = 0
# UART command string for double checking
if sim_mcu == 1:
    uart_cmd_str = "COM" + str(int(mcu_com_addr))
    print(uart_cmd_str)
    mcu_com = rm.open_resource(uart_cmd_str)
    # watch out the form of the address using for connect COM port or GPIB
# write example
# mcu_com.write('abcde')
# read example
# mcu_res = mcu_com.read()

# # infinite loop check the board status
# c_wait_time = 0
# while 1:
#     GPIB_ctrl = sh_main.range((13, 7)).value
#     time.sleep(wait_time)
#     if GPIB_ctrl == 1:
#         # setup resource manager and check what's on bus (COM and )
#         # rm = pyvisa.ResourceManager()
#         list_instr = rm.list_resources()
#         sh_main.range((12, 7)).value = str(list_instr)
#         # think about how to arrange all the device
#         # or just for the reference in python terminal
#         print(list_instr)
#         # print to terminal

#         # load the device name to realted blanks

#         break
#     sh_org_tab.range('A8').value = 'waiting'
#     # control index at the sheet
#     print("waiting GPIB" + str(c_wait_time))
#     c_wait_time = c_wait_time + 1
#     if (c_wait_time == 100):
#         c_wait_time = 0


# ===================== finished the overall initialization

# ===================== main program start

# only x here, different start point in excel for before and after
if before_after == 0:
    y_start = 4
    x_start = 3
elif before_after == 1:
    y_start = 4
    x_start = 11


if dynamic_NO == 1:
    # if using dynamic number
    c_board = 0
    while c_board < dynamic_board:

        # infinite loop check the board status
        while 1:
            # need to change the reference board number or not
            if first_check == 0:
                c_dy_board2 = c_dy_board
                first_check = 1
                print('first time')
            # printing some output from terminal to for debugging
            print('waiting' + str(c_wait_time))
            c_wait_time = c_wait_time + 1
            # time delay is used to prevent excel updating error
            # control the duration reading from board number from excel
            sh_org_tab.range('A8').value = 'waiting'
            # control index at the sheet
            time.sleep(wait_time)
            if c_wait_time > 100:
                c_wait_time = 0

            # capture the dynamic board settings
            c_dy_board = sh_org_tab.range((4, 1)).value
            #  to prevent error create by entering the uncorrect content
            if type(c_dy_board) == None:
                c_dy_board = c_dy_board2
                # correct if there are error from the new input result

            if c_dy_board != c_dy_board2:
                # reset the status control for c_dy_board2 reference
                first_check = 0
                print('out off the loop' + str(c_dy_board) +
                      ' and ' + str(c_dy_board2))
                # print the indicator and escape from the infinite loop
                sh_org_tab.range('A8').value = 'busy'
                # control index at the sheet
                break
            # jump out loop if finished key in number larger than total board
            if c_dy_board > int(total_board):
                break

        if c_dy_board > int(total_board):
            break
        c_item = 0
        # reset counter before each board
        # setup Vin but no calibration => real operation

        # assign board number
        sh_org_tab.range((y_start + int(c_dy_board) - 1,
                         x_start - 1)).value = c_dy_board

        # setup the power supply and COM port SWIRE command at every board operation

        # power supply configuration and turn on
        if sim_real == 1:
            PWR_supply1.chg_out(pre_vin, pwr_max_I, pwr_use_ch, 'on')
            # delay after power on to prevent transient measurement
            time.sleep(wait_time)
            PWR_supply1.change_V(Vin)
            time.sleep(wait_time)

        if sim_mcu == 1:
            # SWIRE command for the maximum output voltage of ELVDD and ELVSS
            # SWIRE default status need to be high
            uart_cmd_str = chr(1) + chr(int(pulse_ELVSS)) + \
                chr(int(pulse_ELVDD))
            print(uart_cmd_str)
            mcu_com.write(uart_cmd_str)
            time.sleep(wait_time)
            # added 2021/11/2, for AVDD pulse
            if pulse_AVDD != 0:
                uart_cmd_str = chr(1) + chr(int(pulse_AVDD)) + \
                    chr(int(pulse_AVDD))
                print(uart_cmd_str)
                mcu_com.write(uart_cmd_str)
                time.sleep(wait_time)

            # input()

        while c_item < RA_out_V:

            # measurement of different channel voltage
            # 1. need to give different UART command to MCU(adjust relay)
            # 2. delay for transition (relay and output voltage)
            # 3. get voltage from meter
            # 4. save the result

            # UART command sent, control by the c_item

            # command sequence should map with the blank in excel sheet to output

            # change the ouput channel connnect to meter

            # power supply read the input current(first item of every board)
            if c_item == 0 and sim_real == 1:
                # when c_item = 0, read the current from power supply
                # read the curent and put to related blank in the sheet
                v_sim = PWR_supply1.read_iout()
                v_sim = v_sim.replace('A', '')
                time.sleep(wait_time)
                # right shift 1 block for the
                # sh_org_tab.range((y_start + int(c_dy_board) - 1, x_start + c_item)).value =
                # no need to update data in the selection, but before the end of loop

            elif c_item != 0:
                # adjust the output channel first
                if sim_mcu == 1:
                    # first c_item is reading iout
                    mcu_com.write(chr(mode) + mcu_cmd_arry[c_item - 1])
                    print(chr(mode) + mcu_cmd_arry[c_item - 1])
                    time.sleep(wait_time)
                    # input()

                # when c_item != 0, read the voltage from related meter
                if sim_real == 1:
                    v_sim = M1_v_in.mea_v()
                    time.sleep(wait_time)
                # MCU UART output only for the output votlage measurement
                # if sim_mcu == 1:
                #     mcu_com.write(str(c_item))
                #     print('mcu out:'+str((c_item)))
                #     time.sleep(wait_time)

                # real data won't need to change v_sim by count
                # v_sim = v_sim + 1
            # update the data to sheet here
            if sim_real == 0:
                v_sim = v_sim + 1
                sh_org_tab.range((y_start + int(c_dy_board) - 1,
                                  x_start + c_item)).value = v_sim
                time.sleep(wait_time)
            if sim_real == 1:
                sh_org_tab.range((y_start + int(c_dy_board) - 1,
                                  x_start + c_item)).value = lo.atof(v_sim)
            c_item = c_item + 1
            # update counter at the end of loop

        # power supply need to turn off here
        # for changing the next board
        if sim_real == 1:
            PWR_supply1.change_V(0)

        c_board = c_board + 1
        # update counter at the end of loop

elif dynamic_NO == 0:
    # if not using dynamic number while loop
    # msg_res = win32api.MessageBox(0, 'press enter after new board is ready', 'Change to the next board')

    c_board = 0
    while c_board < total_board:
        msg_res = win32api.MessageBox(
            0, 'press enter after new board is ready', 'Change to the next board')
        # need to create trigger at each board
        c_item = 0
        # reset counter before each board
        # setup Vin but no calibration => real operation
        # assign board number
        sh_org_tab.range((y_start + int(c_board),
                         x_start - 1)).value = c_board + 1

        # setup the power supply and COM port SWIRE command at every board operation

        # power supply configuration and turn on
        if sim_real == 1:
            PWR_supply1.chg_out(pre_vin, pwr_max_I, pwr_use_ch, 'on')
            # delay after power on to prevent transient measurement
            time.sleep(wait_time)
            PWR_supply1.change_V(Vin)
            time.sleep(wait_time)

        if sim_mcu == 1:
            # SWIRE command for the maximum output voltage of ELVDD and ELVSS
            # SWIRE default status need to be high
            uart_cmd_str = chr(1) + chr(int(pulse_ELVSS)) + \
                chr(int(pulse_ELVDD))
            print(uart_cmd_str)
            mcu_com.write(uart_cmd_str)
            time.sleep(wait_time)
            if pulse_AVDD != 0:
                uart_cmd_str = chr(1) + chr(int(pulse_AVDD)) + \
                    chr(int(pulse_AVDD))
                print(uart_cmd_str)
                mcu_com.write(uart_cmd_str)
                time.sleep(wait_time)

            # input()

        while c_item < RA_out_V:

            # command sequence should map with the blank in excel sheet to output

            if c_item == 0 and sim_real == 1:
                # when c_item = 0, read the current from power supply
                # read the curent and put to related blank in the sheet
                v_sim = PWR_supply1.read_iout()
                v_sim = v_sim.replace('A', '')
                time.sleep(wait_time)
                # right shift 1 block for the
                # sh_org_tab.range((y_start + int(c_dy_board) - 1, x_start + c_item)).value =
                # no need to update data in the selection, but before the end of loop
                # elif sim_real == 0 :
                # no need else if and change the update v_sim at the end of loop

            elif c_item != 0:
                # adjust the output channel first
                if sim_mcu == 1:
                    # first c_item is reading iout
                    mcu_com.write(chr(mode) + mcu_cmd_arry[c_item - 1])
                    print(chr(mode) + mcu_cmd_arry[c_item - 1])
                    time.sleep(wait_time)
                    # input()
                # when c_item != 0, read the voltage from related meter
                if sim_real == 1:
                    v_sim = M1_v_in.mea_v()
                    time.sleep(wait_time)
                # # MCU UART output only for the output votlage measurement
                # if sim_mcu == 1:
                #     mcu_com.write(str(c_item))
                #     print('mcu out:'+str((c_item)))
                #     time.sleep(wait_time)

            if sim_real == 0:
                # update result in simulation mode
                v_sim = v_sim + 1
                sh_org_tab.range((y_start + int(c_board),
                                  x_start + c_item)).value = v_sim
                time.sleep(wait_time)
            if sim_real == 1:
                # update the data to sheet here
                sh_org_tab.range(
                    (y_start + c_board, x_start + c_item)).value = lo.atof(v_sim)
            # real data won't need to change v_sim by count
            # v_sim = v_sim + 1
            c_item = c_item + 1
            # update counter at the end of loop

        # power supply need to turn off here
        # for changing the next board
        if sim_real == 1:
            PWR_supply1.change_V(0)
        wb_res.save(result_book_trace)
        c_board = c_board + 1
        # time.sleep(wait_time)

        # update counter at the end of loop

wb_res.save(result_book_trace)
time.sleep(wait_time)
# comments can also use different layer to achieve better reading (refer to below)
# consideration of saving plan
# maybe better to think about where to sace file during operation,
# because if the file close during operation will lose all the data,
# saving plan for result file need to be take into condieration
wb_res.close()

# plan for using if else separate the simuation or real mode
# better for version control
