# this is the sheet control for the excel file and excel portion of the main_obj

# import for the add-on tool

# import for excel control
import xlwings as xw
# this import is for the VBA function
import win32com.client


control_book_trace = 'c:\\py_gary\\test_excel\\IQ_scan_ctrl.xlsm'
# no place to load the trace from excel or program, define by default
result_book_trace = ''
# result trace unable to load yet

new_file_name = ''
# new file name not loaded yet

# ======== excel application related
# 開啟 Excel 的app
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
# ======== excel application related

excel.Workbooks.Open(
    Filename=control_book_trace)
# use excel app to open the control book, so we can call VBA function


wb = xw.books('Eff_inst.xlsm')
# change the open of xlwings to assign the book had been open from the excel app
# # open control workbook
# wb = xw.Book(control_book_trace)

# define new result workbook
wb_res = xw.Book()
# create reference sheet (for sheet position)
# sh_ref is the index for result sheet
# sh_ref_condition is for testing condition and setting
# all the reference sheet will delete after the program finished
sh_ref = wb_res.sheets.add('ref_sh')
sh_ref_condition = wb_res.sheets.add('ref_sh2')
# delete the extra sheet from new workbook, difference from version
wb_res.sheets('工作表1').delete()

# assign main sheet
sh_main = wb.sheets('main')
# copy the sheets to new book
sh_main.copy(sh_ref)
# assign both sheet to the new sheets in result book
sh_main = wb_res.sheets('main')

# only the instrument control will be still mapped to the original excel
# since inst_ctrl is no needed to copy to the result sheet
sh_inst_ctrl = wb.sheets('inst_ctrl')


# file name from the master excel
new_file_name = str(sh_main.range('B8').value)
# load the trace as string
excel_temp = str(sh_main.range('B9').value)
# program control variable (auto and inst settings)
auto_inst_ctrl = sh_main.range('B11').value
# program exit interrupt variable
program_exit = sh_main.range('B12').value
# verification re-run
re_run_verification = sh_main.range('B13').value

# by loading the start index of each section, no issue for the

# update the result book trace
result_book_trace = excel_temp + new_file_name + '.xlsx'

# save the result book and turn off the control book
wb_res.save(result_book_trace)


# since there will be instrument control, master main sheet won't close
# but the main of the sub program can be close after cpoy and start up
# the operation will be put in the result and test condition will be copy to the
# result book~
# close control book
# wb.close()

# base on output format copied from the control book
# start parameter initialization
# pre- test condition settings
pre_test_en = sh_main.range('C19').value
pre_vin = sh_main.range('C16').value
pre_vin_max = sh_main.range('C17').value
pre_imax = sh_main.range('C18').value
pre_sup_iout = sh_main.range('C20').value

# load the GPIB address for the instrument
# GPIB instrument list (address loading, name feed back)
pwr_supply_addr = sh_main.range('C27').value
meter1_addr = sh_main.range('C28').value
meter2_addr = sh_main.range('C29').value
loader_addr = sh_main.range('C30').value
loader_src_addr = sh_main.range('C31').value
temp_chamber_addr = sh_main.range('C32').value


# other control parameter
# COM port parameter input
mcu_com_addr = sh_main.range('C50').value
# general delay time
wait_time = sh_main.range('C51').value
# control mode for MCU
channel_mode = sh_main.range('C52').value
# pre-increase for efficiency measurement
pre_inc_vin = sh_main.range('C53').value
# the setting for Vin calibration accuracy
vin_diff_set = sh_main.range('C54').value
# the current limit of the VIBAS channel
sw_i2c_select = sh_main.range('C55').value
loader_cal_off1 = sh_main.range('C56').value
loader_cal_off2 = sh_main.range('C57').value
raw_y_position_start = sh_main.range('C58').value
raw_x_position_start = sh_main.range('C59').value

source_meter_channel = sh_main.range('C60').value
# if the channel 1=> EL power, 2=> AVDD, 0=> not use source meter
# when control = 0, all channel used chroma's output mapping
# and this control variable have bigger priority

inst_pwr_ch1 = sh_main.range('C61').value
inst_pwr_ch2 = sh_main.range('C62').value
