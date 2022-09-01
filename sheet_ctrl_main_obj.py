# this is the sheet control for the excel file and excel portion of the main_obj

# import for the add-on tool

# import for excel control
import xlwings as xw
# this import is for the VBA function
import win32com.client

# inport the parameter


control_book_trace = 'c:\\py_gary\\test_excel\\obj_main.xlsm'
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


wb = xw.books('obj_main.xlsm')
# change the open of xlwings to assign the book had been open from the excel app
# # open control workbook
# wb = xw.Book(control_book_trace)

'''
# 220901 generation of result sheet is move to the object
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
#  220901 check if the file setting is mapped to the origin


# copy the sheets to new book
sh_main.copy(sh_ref)
# assign both sheet to the new sheets in result book
sh_main = wb_res.sheets('main')

'''

# assign main sheet
sh_main = wb.sheets('main')

# only the instrument control will be still mapped to the original excel
# since inst_ctrl is no needed to copy to the result sheet
sh_inst_ctrl = wb.sheets('inst_ctrl')

# other way to define sheet:
# this is the format for the efficiency result
ex_sheet_name = 'raw_out'
sh_raw_out = wb.sheets(ex_sheet_name)
# this is the sheet for efficiency testing command
sh_volt_curr_cmd = wb.sheets('V_I_com')
# this is the sheet for I2C command
sh_i2c_cmd = wb.sheets('I2C_ctrl')


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
# file settings; decide which file to load parameters
file_setting = sh_main.range('B14').value

# by loading the start index of each section, no issue for the

# update the result book trace
result_book_trace = excel_temp + new_file_name + '.xlsx'

# # save the result book and turn off the control book
# wb_res.save(result_book_trace)


# since there will be instrument control, master main sheet won't close
# but the main of the sub program can be close after cpoy and start up
# the operation will be put in the result and test condition will be copy to the
# result book~
# close control book
# wb.close()

'''
# 220901: all the parameter loaded function used the other object for definition
# import parameter_load_obj and send the book name needed for the parameter


# insturment parameter loading, to load the instrumenet paramenter
# need to get the index of each item first
# index need to change if adding new control parameter

index_par_pre_con = sh_main.range((3, 9)).value
index_GPIB_inst = sh_main.range((4, 9)).value
index_general_other = sh_main.range((5, 9)).value
index_pwr_inst = sh_main.range((6, 9)).value
index_chroma_inst = sh_main.range((3, 12)).value
index_src_inst = sh_main.range((4, 12)).value
index_meter_inst = sh_main.range((5, 12)).value
index_chamber_inst = sh_main.range((6, 12)).value
index_IQ_scan = sh_main.range((3, 15)).value
index_eff = sh_main.range((4, 15)).value
# index_meter_inst = sh_main.range((5, 15)).value
# index_chamber_inst = sh_main.range((6, 15)).value

# base on output format copied from the control book
# start parameter initialization
# pre- test condition settings
pre_vin = sh_main.range(
    (index_par_pre_con + 1, 3)).value
pre_vin_max = sh_main.range(
    (index_par_pre_con + 2, 3)).value
pre_imax = sh_main.range(
    (index_par_pre_con + 3, 3)).value
pre_test_en = sh_main.range(
    (index_par_pre_con + 4, 3)).value
pre_sup_iout = sh_main.range(
    (index_par_pre_con + 5, 3)).value

# load the GPIB address for the instrument
# GPIB instrument list (address loading, name feed back)
pwr_supply_addr = sh_main.range(
    (index_GPIB_inst + 1, 3)).value
# met1 usually for voltage
meter1_addr = sh_main.range(
    (index_GPIB_inst + 2, 3)).value
# met2 usually for current
meter2_addr = sh_main.range(
    (index_GPIB_inst + 3, 3)).value
loader_addr = sh_main.range(
    (index_GPIB_inst + 4, 3)).value
loader_src_addr = sh_main.range(
    (index_GPIB_inst + 5, 3)).value
temp_chamber_addr = sh_main.range(
    (index_GPIB_inst + 6, 3)).value

# initialization for all the object, based on the input parameter of the index

# parameter setting for the power supply
pwr_vset = sh_main.range((index_pwr_inst + 1, 3)).value
pwr_iset = sh_main.range((index_pwr_inst + 2, 3)).value
pwr_act_ch = sh_main.range(
    (index_pwr_inst + 3, 3)).value
pwr_ini_state = sh_main.range(
    (index_pwr_inst + 4, 3)).value
relay0_ch = sh_main.range((index_pwr_inst + 5, 3)).value
relay6_ch = sh_main.range((index_pwr_inst + 6, 3)).value
relay7_ch = sh_main.range((index_pwr_inst + 7, 3)).value
# pre-increase for efficiency measurement
pre_inc_vin = sh_main.range(
    (index_pwr_inst + 8, 3)).value
# the setting for Vin calibration accuracy
vin_diff_set = sh_main.range(
    (index_pwr_inst + 9, 3)).value

# parameter setting for the chroma loader
loader_act_ch = sh_main.range(
    (index_chroma_inst + 1, 3)).value
loader_ini_mode = sh_main.range(
    (index_chroma_inst + 2, 3)).value
loader_cal_ELch = sh_main.range(
    (index_chroma_inst + 3, 3)).value
loader_cal_VCIch = sh_main.range(
    (index_chroma_inst + 4, 3)).value
loader_ELch = sh_main.range(
    (index_chroma_inst + 5, 3)).value
loader_ini_state = sh_main.range(
    (index_chroma_inst + 6, 3)).value
loader_VCIch = sh_main.range(
    (index_chroma_inst + 7, 3)).value
loader_cal_mode = sh_main.range(
    (index_chroma_inst + 8, 3)).value

# parameter setting for source meter
src_vset = sh_main.range((index_src_inst + 1, 3)).value
src_iset = sh_main.range((index_src_inst + 2, 3)).value
src_ini_state = sh_main.range(
    (index_src_inst + 3, 3)).value
src_ini_type = sh_main.range(
    (index_src_inst + 4, 3)).value
src_clamp_ini = sh_main.range(
    (index_src_inst + 5, 3)).value

# parameter setting for meter
met_v_res = sh_main.range(
    (index_meter_inst + 1, 3)).value
met_v_max = sh_main.range(
    (index_meter_inst + 2, 3)).value
met_i_res = sh_main.range(
    (index_meter_inst + 3, 3)).value
met_i_max = sh_main.range(
    (index_meter_inst + 4, 3)).value

# parameter setting for chamber

cham_tset_ini = sh_main.range(
    (index_chamber_inst + 1, 3)).value
cham_ini_state = sh_main.range(
    (index_chamber_inst + 2, 3)).value
cham_l_limt = sh_main.range(
    (index_chamber_inst + 3, 3)).value
cham_h_limt = sh_main.range(
    (index_chamber_inst + 4, 3)).value
cham_hyst = sh_main.range(
    (index_chamber_inst + 5, 3)).value

# other control parameter
# COM port parameter input
mcu_com_addr = sh_main.range(
    index_general_other + 1, 3).value
# general delay time
wait_time = sh_main.range(
    index_general_other + 2, 3).value
# the start point for the raw_out index
raw_y_position_start = sh_main.range(
    index_general_other + 3, 3).value
raw_x_position_start = sh_main.range(
    index_general_other + 4, 3).value
sheet_off_finished = sh_main.range(
    index_general_other + 5, 3).value
# plot pause control (1 is enable, 0 is disable)
en_plot_waring = sh_main.range(
    index_general_other + 6, 3).value
en_fully_auto = sh_main.range(
    index_general_other + 7, 3).value
en_start_up_check = sh_main.range(
    index_general_other + 8, 3).value

# verification item: IQ parameter
ISD_range = sh_main.range(
    index_IQ_scan + 1, 3).value

# verification item: eff control parameter
channel_mode = sh_main.range(index_eff + 1, 3).value
# SWIRE or I2C selected setting
sw_i2c_select = sh_main.range(index_eff + 2, 3).value
# if the channel 1=> EL power, 2=> AVDD, 0=> not use source meter
# when control = 0, all channel used chroma's output mapping
source_meter_channel = sh_main.range(
    index_eff + 3, 3).value

print('end of the parameter loaded')

'''
