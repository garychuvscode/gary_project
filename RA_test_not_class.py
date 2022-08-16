import xlwings as xw
# for excel related function
import pyvisa
# for UART and GPIB related function
import time
# for delay related function
import win32api
# for message box and others
from win32con import MB_SYSTEMMODAL
# other function for message

# response = win32api.MessageBox(0, "Did you hear the Buzzer?", "Buzzer Test", 2, MB_SYSTEMMODAL)
# response2 = win32api.MessageBox(0, 'hello', 'title')


# ======== GPIB and UART initialization
# consider to put the control of instrument into sub program in other file

#rm = pyvisa.ResourceManager()
#list_instr = rm.list_resources()
# print(list_instr)
#('ASRL1::INSTR', 'ASRL2::INSTR', 'GPIB0::12::INSTR')
#inst = rm.open_resource('GPIB0::1::INSTR')
#inst2 = rm.open_resource('GPIB0::22::INSTR')


# print(inst.query("*IDN?"))

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


# open control workbook
wb = xw.Book('c:\\py_gary\\test_excel\\RA_control.xlsx')
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
wb_res.save('c:\\py_gary\\test_excel\\RA.xlsx')
wb.close()
# close control sheets

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

# message box return result (most time is int)
msg_res = 0

# ===================== finished the overall initialization

# GPIB initialization

# infinite loop check the board status
c_wait_time = 0
while 1:
    GPIB_ctrl = sh_main.range((13, 7)).value
    time.sleep(wait_time)
    if GPIB_ctrl == 1:
        # setup resource manager and check what's on bus (COM and )
        rm = pyvisa.ResourceManager()
        list_instr = rm.list_resources()
        sh_main.range((12, 7)).value = str(list_instr)
        print(list_instr)
        break

    print("waiting GPIB" + str(c_wait_time))
    c_wait_time = c_wait_time + 1
    if (c_wait_time == 100):
        c_wait_time = 0

dc_source = rm.open_resource('GPIB0::1::INSTR')
meter_v = rm.open_resource('GPIB0::22::INSTR')

print(dc_source.query("*IDN?"))
print(meter_v.query("*IDN?"))


# DC source initialization string set
# default settings
vset = 3.7
iset = 1
act_ch = 3
# on or off oh channel (default off)
dc_sw = 'off'

# assume to turn the output on from somethings
dc_sw = 'on'

print('Vset: ' + str(vset) + ' Iset: ' + str(iset) + ' CH: ' + str(act_ch))
# show the vset and iset value from the window

# turn off the channel
# need to add dc_sw selection before output the command string, or it will turn off
# using control to recover to off when finished command
cmd_str_out_sw = "OUT" + str(act_ch) + ":" + str(dc_sw)
# using normal(independent mode, not serial)
cmd_str_out_mode = "OUT: normal"
# setting string for the V and I
cmd_str_V = ("PROG:VSET" + str(act_ch) + ":" + str(vset))
cmd_str_I = ("PROG:ISET" + str(act_ch) + ":" + str(iset))

# import example for the GPIB operating wirte and read
# my_instrument.write('*IDN?')
# print(my_instrument.read())

# the equivalent command is:

# my_instrument.query("*IDN?") => which include write and one read follow
# print(meter_v.query("*IDN?")) => will print the device ID out
# meter_ID = meter_v.query("*IDN?") => will assign the device ID into variable meter_ID

# ----DC cource control
# writing the control command for DC source

dc_source.write(cmd_str_out_mode)
dc_source.write(cmd_str_V)
dc_source.write(cmd_str_I)
dc_source.write(cmd_str_out_sw)

# add other part to changing command
vset = 2
iset = 2

cmd_str_V = ("PROG:VSET" + str(act_ch) + ":" + str(vset))
cmd_str_I = ("PROG:ISET" + str(act_ch) + ":" + str(iset))

dc_source.write(cmd_str_V)
dc_source.write(cmd_str_I)
# when the source is already on, need to have the turn on command to update the final
# setting to output, to prevent wrong behavior of the power supply
# (only change the output voltage but not the current)
# use another command to update at the same time
dc_source.write(cmd_str_out_sw)


# ----DC cource control end


cmd_str_mea_v = 'MEAS:VOLT:DC? 7,0.0001'
cmd_str_mea_i = 'MEAS:CURR:DC? 0.1,0.000001'
# delay for the voltage and current ready
time.sleep(wait_time)

# ---- meter read in realinstrument control
vres_test = meter_v.query(cmd_str_mea_v)
ires_test = meter_v.query(cmd_str_mea_i)

print(vres_test)
print(ires_test)
# ---- meter read in realinstrument control end

# meter initialization string set


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
            time.sleep(wait_time)
            if c_wait_time > 100:
                c_wait_time = 0

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
                break
            if c_dy_board > int(total_board):
                break

        if c_dy_board > int(total_board):
            break
        c_item = 0
        # reset counter before each board
        # setup Vin but no calibration => real operation
        while c_item < RA_out_V:

            # assign board number
            sh_org_tab.range((y_start + int(c_dy_board) - 1, x_start - 1)
                             ).value = c_dy_board
            sh_org_tab.range(
                (y_start + int(c_dy_board) - 1, x_start + c_item)).value = v_sim
            # real data won't need to change v_sim by count
            v_sim = v_sim + 1
            c_item = c_item + 1
            # update counter at the end of loop

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
        while c_item < RA_out_V:
            sh_org_tab.range((y_start + c_board, x_start - 1)
                             ).value = c_board + 1
            sh_org_tab.range(
                (y_start + c_board, x_start + c_item)).value = v_sim
            # real data won't need to change v_sim by count
            v_sim = v_sim + 1
            c_item = c_item + 1
            # update counter at the end of loop

        c_board = c_board + 1
        time.sleep(wait_time)

        # update counter at the end of loop

wb_res.save('c:\\py_gary\\test_excel\\RA.xlsx')
wb_res.close()
