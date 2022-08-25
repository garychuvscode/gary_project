# sheet_control
# this file is used to initializtion the control book and gen the result book

# need to check if the import function can load the sheet variable directly to the main py file

import win32com.client
# this import is for the VBA function

from glob import glob
import webbrowser
from idna import check_nfc
from openpyxl import Workbook
import xlwings as xw
# import the excel control package

import numpy as np

# test mode selection for sheet control test
test_mode = 1
# test mode setting index:
# 0 => test before 220516
# 1 => inst control related testing


# extra_name = '0'
# extra name for i2c or sw information on file name

excel_temp = 0
sheet_temp = 0
# temp used variable

sheet_arry = np.full([200], None)
# sheet_arry = np.zeros(100)
file_array = np.full([100], None)
# file array used to save all the file name for ixnde

excel_trace = 0
# trace control variable

file_count = 3
# testing file amount for the file build function

sub_sh_count = 0

# 220824 eff_done related control
eff_done_sh = 0

# ======== excel book and sheet operation
# control book loading => new book creation => default parameter loading

control_book_trace = 'c:\\py_gary\\test_excel\\Eff_inst.xlsm'
# no place to load the trace from excel or program, define by default
result_book_trace = ''
# result trace unable to load yet

new_file_name = ''
# new file name not loaded yet

# open the file from teh excel app

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

# take the varaiable out for other place reference
# by using the golbal variable, these variable can be used at other part of program
# or be used in other file import control sheet

# ========== first time, operate the maain program and build global variable

# define new result workbook

# 220324 => for the stuff related to build new file, move to the build file sub-program
# wb_res = xw.Book()
# # create reference sheet (for sheet position)
# sh_ref = wb_res.sheets.add('ref_sh')
# # delete the extra sheet from new workbook, difference from version
# wb_res.sheets('工作表1').delete()


# define the sheets in control book
result_sheet_name = 'raw_out'
sh_main = wb.sheets('main')
sh_org_tab = wb.sheets('V_I_com')
sh_org_tab2 = wb.sheets(result_sheet_name)
sh_org_tab3 = wb.sheets('I2C_ctrl')
sh_inst_ctrl = wb.sheets('inst_ctrl')
# instrument control only keep in wb, not to change if the wb_res close


# # 220324 => move to build file sub-program
# # copy the sheets to new book
# sh_main.copy(sh_ref)
# sh_org_tab.copy(sh_ref)
# sh_org_tab2.copy(sh_ref)
# sh_org_tab3.copy(sh_ref)
wb_res = xw.Book()
wb_res.close()

# assign both sheet to the new sheets in result book
# sh_main = wb_res.sheets('main')
# sh_org_tab = wb_res.sheets('V_I_com')
# sh_org_tab2 = wb_res.sheets(result_sheet_name)
# sh_org_tab3 = wb_res.sheets('I2C_ctrl')
# sh_main = sh_main.copy(sh_ref)
# sh_org_tab = sh_org_tab.copy(sh_ref)
# delete reference after copied

# 20220112: here is only loading the sheet with command
# the result initialization will be update after the parameter
# are loaded into python
# the result only copy at the end of she control

# old settings for delete the reference sheet here
# 20220112 => moe to the end of sheet control because it may add many sheet
# based on the conrtrol setting loaded from the excel
# sh_ref.delete()

excel_trace = str(sh_main.range('B9').value)
# load the trace as string
new_file_name = str(sh_main.range('B8').value)

# 220518 add the control variable for the inst and auto selection
# 0 => only auto; 1 => only inst; 2 => auto + inst
inst_auto_selection = sh_main.range('B11').value
# prograam keep going if program_exit > 0
program_exit = sh_main.range('B12').value

# # update the result book trace
# result_book_trace = excel_trace + new_file_name + str(extra_name) + '.xlsx'
result_book_trace = ''
full_result_name = ''

# 220824 reset the eff_done_sh
sh_main.range('C13').value = 0
# 220824 add the turn off control of finished sheet
sheet_off_finished = sh_main.range('C63').value
# plot pause control (1 is enable, 0 is disable)
en_plot_waring = sh_main.range('C64').value
en_fully_auto = sh_main.range('C65').value
en_start_up_check = sh_main.range('C66').value
en_chamber_mea = sh_main.range('C67').value
chamber_default_tset = sh_main.range('C68').value
chamber_H_limt = sh_main.range('C69').value
chamber_L_limt = sh_main.range('C70').value
chamber_error = sh_main.range('C71').value


# # save the result book and turn off the control book
# # 220324 move to build file
# wb_res.save(result_book_trace)
# # wb.close()
# # close control books

# 20220502 instrument control parameter loading
# use update sub program for continuous update the parameter from wb
# position index (y-pos, x-pos)
insref_pwr_y = sh_inst_ctrl.range('J5').value
insref_load_y = sh_inst_ctrl.range('J6').value
insref_met1_y = sh_inst_ctrl.range('J7').value
insref_met2_y = sh_inst_ctrl.range('J8').value
insref_src_y = sh_inst_ctrl.range('J9').value

insref_pwr_x = sh_inst_ctrl.range('K5').value
insref_load_x = sh_inst_ctrl.range('K6').value
insref_met1_x = sh_inst_ctrl.range('K7').value
insref_met2_x = sh_inst_ctrl.range('K8').value
insref_src_x = sh_inst_ctrl.range('K9').value

# insctl_refresh = sh_inst_ctrl.range('N4').value

# variable to get the status update from main program

pwr_connection_status = 0

pwr_v_ch1_status = 0
pwr_i_ch1_status = 0
pwr_o_ch1_status = 0
pwr_auto_ch1_status = 1

pwr_v_ch2_status = 0
pwr_i_ch2_status = 0
pwr_o_ch2_status = 0
pwr_auto_ch2_status = 1

pwr_v_ch3_status = 0
pwr_i_ch3_status = 0
pwr_o_ch3_status = 0
pwr_auto_ch3_status = 1


load_connection_status = 0

load_i_ch1_status = 0
load_m_ch1_status = 0
load_o_ch1_status = 0
load_auto_ch1_status = 1

load_i_ch2_status = 0
load_m_ch2_status = 0
load_o_ch2_status = 0
load_auto_ch2_status = 1

load_i_ch3_status = 0
load_m_ch3_status = 0
load_o_ch3_status = 0
load_auto_ch3_status = 1

load_i_ch4_status = 0
load_m_ch4_status = 0
load_o_ch4_status = 0
load_auto_ch4_status = 1

met1_connection_status = 0
met1_mode_status = 0
met1_level_status = 0
met1_v_mea_status = 0
met1_i_mea_status = 0

met2_connection_status = 0
met2_mode_status = 0
met2_level_status = 0
met2_v_mea_status = 0
met2_i_mea_status = 0

src_connection_status = 0
src_mode_status = 0
src_clamp_status = 0
src_level_status = 0
src_v_mea_status = 0
src_i_mea_status = 0
src_o_status = 0


# inssts => instrument status related parameter (blue)
# status output indexing (y-pos and x-pos)
inssts_pwr_connection_y = int(insref_pwr_y)
inssts_pwr_connection_x = int(insref_pwr_x) + 1

inssts_pwr_refresh_y = int(insref_pwr_y) + 1
inssts_pwr_refresh_x = int(insref_pwr_x) + 2

inssts_pwr_serial_y = int(insref_pwr_y)
inssts_pwr_serial_x = int(insref_pwr_x) + 5

inssts_pwr_calibration_y = int(insref_pwr_y) + 1
inssts_pwr_calibration_x = int(insref_pwr_x) + 5

inssts_pwr_vset_ch1_y = int(insref_pwr_y) + 4
inssts_pwr_vset_ch1_x = int(insref_pwr_x) + 1
inssts_pwr_iset_ch1_y = int(insref_pwr_y) + 6
inssts_pwr_iset_ch1_x = int(insref_pwr_x) + 1
inssts_pwr_outs_ch1_y = int(insref_pwr_y) + 8
inssts_pwr_outs_ch1_x = int(insref_pwr_x) + 1
inssts_pwr_auto_ch1_y = int(insref_pwr_y) + 2
inssts_pwr_auto_ch1_x = int(insref_pwr_x) + 1

inssts_pwr_vset_ch2_y = int(insref_pwr_y) + 4
inssts_pwr_vset_ch2_x = int(insref_pwr_x) + 3
inssts_pwr_iset_ch2_y = int(insref_pwr_y) + 6
inssts_pwr_iset_ch2_x = int(insref_pwr_x) + 3
inssts_pwr_outs_ch2_y = int(insref_pwr_y) + 8
inssts_pwr_outs_ch2_x = int(insref_pwr_x) + 3
inssts_pwr_auto_ch2_y = int(insref_pwr_y) + 2
inssts_pwr_auto_ch2_x = int(insref_pwr_x) + 3

inssts_pwr_vset_ch3_y = int(insref_pwr_y) + 4
inssts_pwr_vset_ch3_x = int(insref_pwr_x) + 5
inssts_pwr_iset_ch3_y = int(insref_pwr_y) + 6
inssts_pwr_iset_ch3_x = int(insref_pwr_x) + 5
inssts_pwr_outs_ch3_y = int(insref_pwr_y) + 8
inssts_pwr_outs_ch3_x = int(insref_pwr_x) + 5
inssts_pwr_auto_ch3_y = int(insref_pwr_y) + 2
inssts_pwr_auto_ch3_x = int(insref_pwr_x) + 5

# position of index for channel status (for status color change)
inssts_pwr_pos_ch1_y = int(insref_pwr_y) + 2
inssts_pwr_pos_ch1_x = int(insref_pwr_x) + 0
inssts_pwr_pos_ch2_y = int(insref_pwr_y) + 2
inssts_pwr_pos_ch2_x = int(insref_pwr_x) + 2
inssts_pwr_pos_ch3_y = int(insref_pwr_y) + 2
inssts_pwr_pos_ch3_x = int(insref_pwr_x) + 4

# DC loader status output index
inssts_load_connection_y = int(insref_load_y)
inssts_load_connection_x = int(insref_load_x) + 1

inssts_load_refresh_y = int(insref_load_y) + 1
inssts_load_refresh_x = int(insref_load_x) + 2

inssts_load_iset_ch1_y = int(insref_load_y) + 4
inssts_load_iset_ch1_x = int(insref_load_x) + 1
inssts_load_mset_ch1_y = int(insref_load_y) + 6
inssts_load_mset_ch1_x = int(insref_load_x) + 1
inssts_load_outs_ch1_y = int(insref_load_y) + 8
inssts_load_outs_ch1_x = int(insref_load_x) + 1
inssts_load_auto_ch1_y = int(insref_load_y) + 2
inssts_load_auto_ch1_x = int(insref_load_x) + 1

inssts_load_iset_ch2_y = int(insref_load_y) + 4
inssts_load_iset_ch2_x = int(insref_load_x) + 3
inssts_load_mset_ch2_y = int(insref_load_y) + 6
inssts_load_mset_ch2_x = int(insref_load_x) + 3
inssts_load_outs_ch2_y = int(insref_load_y) + 8
inssts_load_outs_ch2_x = int(insref_load_x) + 3
inssts_load_auto_ch2_y = int(insref_load_y) + 2
inssts_load_auto_ch2_x = int(insref_load_x) + 3

inssts_load_iset_ch3_y = int(insref_load_y) + 4
inssts_load_iset_ch3_x = int(insref_load_x) + 5
inssts_load_mset_ch3_y = int(insref_load_y) + 6
inssts_load_mset_ch3_x = int(insref_load_x) + 5
inssts_load_outs_ch3_y = int(insref_load_y) + 8
inssts_load_outs_ch3_x = int(insref_load_x) + 5
inssts_load_auto_ch3_y = int(insref_load_y) + 2
inssts_load_auto_ch3_x = int(insref_load_x) + 5

inssts_load_iset_ch4_y = int(insref_load_y) + 4
inssts_load_iset_ch4_x = int(insref_load_x) + 7
inssts_load_mset_ch4_y = int(insref_load_y) + 6
inssts_load_mset_ch4_x = int(insref_load_x) + 7
inssts_load_outs_ch4_y = int(insref_load_y) + 8
inssts_load_outs_ch4_x = int(insref_load_x) + 7
inssts_load_auto_ch4_y = int(insref_load_y) + 2
inssts_load_auto_ch4_x = int(insref_load_x) + 7

# muti-meter status output index
inssts_met1_connection_y = int(insref_met1_y)
inssts_met1_connection_x = int(insref_met1_x) + 1
inssts_met1_refresh_y = int(insref_met1_y) + 7
inssts_met1_refresh_x = int(insref_met1_x) + 1
inssts_met1_mset_y = int(insref_met1_y) + 3
inssts_met1_mset_x = int(insref_met1_x) + 1
inssts_met1_leve_y = int(insref_met1_y) + 5
inssts_met1_leve_x = int(insref_met1_x) + 1
inssts_met1_meav_y = int(insref_met1_y) + 7
inssts_met1_meav_x = int(insref_met1_x) + 0
inssts_met1_meai_y = int(insref_met1_y) + 9
inssts_met1_meai_x = int(insref_met1_x) + 0

inssts_met2_connection_y = int(insref_met2_y)
inssts_met2_connection_x = int(insref_met2_x) + 1
inssts_met2_refresh_y = int(insref_met2_y) + 7
inssts_met2_refresh_x = int(insref_met2_x) + 1
inssts_met2_mset_y = int(insref_met2_y) + 3
inssts_met2_mset_x = int(insref_met2_x) + 1
inssts_met2_leve_y = int(insref_met2_y) + 5
inssts_met2_leve_x = int(insref_met2_x) + 1
inssts_met2_meav_y = int(insref_met2_y) + 7
inssts_met2_meav_x = int(insref_met2_x) + 0
inssts_met2_meai_y = int(insref_met2_y) + 9
inssts_met2_meai_x = int(insref_met2_x) + 0

# source meter status output index

inssts_src_connection_y = int(insref_src_y)
inssts_src_connection_x = int(insref_src_x) + 1
inssts_src_refresh_y = int(insref_src_y) + 9
inssts_src_refresh_x = int(insref_src_x) + 1
inssts_src_cset_y = int(insref_src_y) + 3
inssts_src_cset_x = int(insref_src_x) + 1
inssts_src_mset_y = int(insref_src_y) + 5
inssts_src_mset_x = int(insref_src_x) + 1
inssts_src_leve_y = int(insref_src_y) + 7
inssts_src_leve_x = int(insref_src_x) + 1
inssts_src_meav_y = int(insref_src_y) + 9
inssts_src_meav_x = int(insref_src_x) + 0
inssts_src_meai_y = int(insref_src_y) + 11
inssts_src_meai_x = int(insref_src_x) + 0
inssts_src_outs_y = int(insref_src_y) + 13
inssts_src_outs_x = int(insref_src_x) + 1


# finished the output index settings
# since it will be fixed after table reference cell is set, only need to load once

# insctl => instrument control related parameter (green)
# input parameter loading

# DC power supply table
insctl_pwr_refresh = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 1, int(insref_pwr_x) + 1)).value
insctl_pwr_serial = sh_inst_ctrl.range(
    (int(insref_pwr_y), int(insref_pwr_x) + 4)).value
insctl_pwr_calibration = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 1, int(insref_pwr_x) + 4)).value
insctl_pwr_vset_ch1 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 4, int(insref_pwr_x) + 0)).value
insctl_pwr_iset_ch1 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 6, int(insref_pwr_x) + 0)).value
insctl_pwr_outs_ch1 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 8, int(insref_pwr_x) + 0)).value
insctl_pwr_vset_ch2 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 4, int(insref_pwr_x) + 2)).value
insctl_pwr_iset_ch2 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 6, int(insref_pwr_x) + 2)).value
insctl_pwr_outs_ch2 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 8, int(insref_pwr_x) + 2)).value
insctl_pwr_vset_ch3 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 4, int(insref_pwr_x) + 4)).value
insctl_pwr_iset_ch3 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 6, int(insref_pwr_x) + 4)).value
insctl_pwr_outs_ch3 = sh_inst_ctrl.range(
    (int(insref_pwr_y) + 8, int(insref_pwr_x) + 4)).value

# DC loader table
insctl_load_refresh = sh_inst_ctrl.range(
    (int(insref_load_y) + 1, int(insref_load_x) + 1)).value
insctl_load_iset_ch1 = sh_inst_ctrl.range(
    (int(insref_load_y) + 4, int(insref_load_x) + 0)).value
insctl_load_mset_ch1 = sh_inst_ctrl.range(
    (int(insref_load_y) + 6, int(insref_load_x) + 0)).value
insctl_load_outs_ch1 = sh_inst_ctrl.range(
    (int(insref_load_y) + 8, int(insref_load_x) + 0)).value
insctl_load_iset_ch2 = sh_inst_ctrl.range(
    (int(insref_load_y) + 4, int(insref_load_x) + 2)).value
insctl_load_mset_ch2 = sh_inst_ctrl.range(
    (int(insref_load_y) + 6, int(insref_load_x) + 2)).value
insctl_load_outs_ch2 = sh_inst_ctrl.range(
    (int(insref_load_y) + 8, int(insref_load_x) + 2)).value
insctl_load_iset_ch3 = sh_inst_ctrl.range(
    (int(insref_load_y) + 4, int(insref_load_x) + 4)).value
insctl_load_mset_ch3 = sh_inst_ctrl.range(
    (int(insref_load_y) + 6, int(insref_load_x) + 4)).value
insctl_load_outs_ch3 = sh_inst_ctrl.range(
    (int(insref_load_y) + 8, int(insref_load_x) + 4)).value
insctl_load_iset_ch4 = sh_inst_ctrl.range(
    (int(insref_load_y) + 4, int(insref_load_x) + 6)).value
insctl_load_mset_ch4 = sh_inst_ctrl.range(
    (int(insref_load_y) + 6, int(insref_load_x) + 6)).value
insctl_load_outs_ch4 = sh_inst_ctrl.range(
    (int(insref_load_y) + 8, int(insref_load_x) + 6)).value

# multi-meter talbe

insctl_met1_refresh = sh_inst_ctrl.range(
    (int(insref_met1_y) + 1, int(insref_met1_x) + 1)).value
insctl_met1_mset = sh_inst_ctrl.range(
    (int(insref_met1_y) + 3, int(insref_met1_x) + 0)).value
# the measurement level of the meter
insctl_met1_leve = sh_inst_ctrl.range(
    (int(insref_met1_y) + 5, int(insref_met1_x) + 0)).value

insctl_met2_refresh = sh_inst_ctrl.range(
    (int(insref_met2_y) + 1, int(insref_met2_x) + 1)).value
insctl_met2_mset = sh_inst_ctrl.range(
    (int(insref_met2_y) + 3, int(insref_met2_x) + 0)).value
# the measurement level of the meter
insctl_met2_leve = sh_inst_ctrl.range(
    (int(insref_met2_y) + 5, int(insref_met2_x) + 0)).value

# source-meter talbe

insctl_src_refresh = sh_inst_ctrl.range(
    (int(insref_src_y) + 1, int(insref_src_x) + 1)).value
# setting of the clamp parameter
insctl_src_cset = sh_inst_ctrl.range(
    (int(insref_src_y) + 3, int(insref_src_x) + 0)).value
insctl_src_mset = sh_inst_ctrl.range(
    (int(insref_src_y) + 5, int(insref_src_x) + 0)).value
# the setting level of the source meter
insctl_src_leve = sh_inst_ctrl.range(
    (int(insref_src_y) + 7, int(insref_src_x) + 0)).value
# source meter ona and off control status
insctl_src_outs = sh_inst_ctrl.range(
    (int(insref_src_y) + 13, int(insref_src_x) + 0)).value


# base on output format copied from the control book
# start parameter initialization
# pre- test condition settings
pre_test_en = sh_main.range('C19').value
pre_vin = sh_main.range('C16').value
pre_vin_max = sh_main.range('C17').value
pre_imax = sh_main.range('C18').value
pre_sup_iout = sh_main.range('C20').value

# GPIB instrument list (address loading, name feed back)
pwr_sup_addr = sh_main.range('C27').value
met_vdd_addr = sh_main.range('C28').value
met_vss_addr = sh_main.range('C29').value
loa_dts_addr = sh_main.range('C30').value
loa_src_addr = sh_main.range('C31').value
temp_chamber_addr = sh_main.range('C32').value

# optional control parameter for instrument
v_clamp_load = sh_main.range('C37').value
# this is for source meter
met_vin_rang = sh_main.range('C38').value
# maximum V for meter
met_iin_rang = sh_main.range('C39').value
# maximum I for meter
inst_pwr_ch3 = sh_main.range('C40').value
# resolution of meter
met_vin_res = sh_main.range('C41').value
met_iin_res = sh_main.range('C42').value
# setting of loader
loa_ch_set = sh_main.range('C43').value
# load channel for EL power
loa_mod_set = sh_main.range('C44').value
loa_ch2_set = sh_main.range('C45').value
# load channel for AVDD
loader_cal_mode = sh_main.range('C46').value
# source type for the source meter
loader_source_type = sh_main.range('C47').value


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


# parameter needed from the result sheet
# setting for the VBIAS, VIN and current, loaded the counter
# into the python

# vbias_en = sh_org_tab.range('H1').value
# vbias_pre = sh_org_tab.range('H2').value
# vin_pre = sh_org_tab.range('H3').value
# load_pre = sh_org_tab.range('H4').value

# counteer is usually use c_ in opening
c_avdd_load = sh_org_tab.range('D1').value
c_vin = sh_org_tab.range('B1').value
c_iload = sh_org_tab.range('C1').value
c_pulse = sh_org_tab.range('E1').value
c_i2c = sh_org_tab3.range('B1').value
c_i2c_g = sh_org_tab3.range('D1').value
c_avdd_single = sh_org_tab.range('G1').value
c_avdd_pulse = sh_org_tab.range('H1').value
c_tempature = sh_org_tab.range('I1').value
# avdd single is single channel current setting for AVDD (1-channel testing)
# using c_pulse or c_i2c is depend on the sw_i2c_select


# # move to build file sub-prog
# # used array to setup the result sheet
# # res_sheet_array = np.zeros(100)
# # this method not working, skip this time

# if sw_i2c_select == 1:

#     # here is to generate the sheet for each VBIAS settings
#     c_sheet_copy = c_avdd_load
#     x_sheet_copy = 0
#     sh_temp = sh_org_tab2
#     while x_sheet_copy < c_sheet_copy:
#         # if x_sheet_copy == 0:
#         sh_temp.copy(sh_ref)
#         sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
#         excel_temp = str(sh_org_tab.range(3 + x_sheet_copy, 4).value)
#         sheet_temp = 'I_AVDD=' + excel_temp + 'A'
#         # save the sheet name into the array for loading
#         sheet_arry[2 * x_sheet_copy] = sheet_temp
#         sh_org_tab2.name = sheet_temp
#         # also assign the value to the VBIAS element
#         # this sheet need to have example(general format)
#         sh_org_tab2.range(21, 3).value = sh_org_tab.range(
#             3 + x_sheet_copy, 4).value
#         # add another sheet for the raw data of each AVDD current
#         sheet_temp = 'raw_I_AVDD=' + excel_temp + 'A'
#         # raw data sheet no need example format, can use empty sheet
#         sheet_arry[2 * x_sheet_copy + 1] = sheet_temp
#         wb_res.sheets.add(sheet_temp)

#         x_sheet_copy = x_sheet_copy + 1

#     sh_temp.delete()
# else:
#     sh_org_tab2 = wb_res.sheets(result_sheet_name)
#     # even there are no other bias settings, also need to set the c_bias
#     # so the loop of measurement can start
#     c_bias = 1

# # 20221012 new sheet delete place is here
# sh_ref.delete()


# wb_res.save(result_book_trace)
# # sub program needed for the contol book initialization

def para_update_pwr():
    # no need to input the parameter, check all the items based on the refrech time setting for each device
    # settings for table no need to use global, because re-load from excel every time,
    # but the status need to be global, since it will save the previous status
    # need to separate all the different instrument (there are different refresh rate)

    # call for parameter updatae
    # DC power supply table
    global insctl_pwr_refresh
    global insctl_pwr_serial
    global insctl_pwr_calibration
    global insctl_pwr_vset_ch1
    global insctl_pwr_iset_ch1
    global insctl_pwr_outs_ch1
    global insctl_pwr_vset_ch2
    global insctl_pwr_iset_ch2
    global insctl_pwr_outs_ch2
    global insctl_pwr_vset_ch3
    global insctl_pwr_iset_ch3
    global insctl_pwr_outs_ch3

    # DC power supply table
    insctl_pwr_refresh = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 1, int(insref_pwr_x) + 1)).value
    insctl_pwr_serial = sh_inst_ctrl.range(
        (int(insref_pwr_y), int(insref_pwr_x) + 4)).value
    insctl_pwr_calibration = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 1, int(insref_pwr_x) + 4)).value
    insctl_pwr_vset_ch1 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 4, int(insref_pwr_x) + 0)).value
    insctl_pwr_iset_ch1 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 6, int(insref_pwr_x) + 0)).value
    insctl_pwr_outs_ch1 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 8, int(insref_pwr_x) + 0)).value
    insctl_pwr_vset_ch2 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 4, int(insref_pwr_x) + 2)).value
    insctl_pwr_iset_ch2 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 6, int(insref_pwr_x) + 2)).value
    insctl_pwr_outs_ch2 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 8, int(insref_pwr_x) + 2)).value
    insctl_pwr_vset_ch3 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 4, int(insref_pwr_x) + 4)).value
    insctl_pwr_iset_ch3 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 6, int(insref_pwr_x) + 4)).value
    insctl_pwr_outs_ch3 = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 8, int(insref_pwr_x) + 4)).value

    print('para_update_pwr')
    print('insctl_pwr_refresh = ' + str(insctl_pwr_refresh))
    print('insctl_pwr_serial = ' + str(insctl_pwr_serial))
    print('insctl_pwr_calibration = ' + str(insctl_pwr_calibration))
    print('insctl_pwr_vset_ch1 = ' + str(insctl_pwr_vset_ch1))
    print('insctl_pwr_iset_ch1 = ' + str(insctl_pwr_iset_ch1))
    print('insctl_pwr_outs_ch1 = ' + str(insctl_pwr_outs_ch1))
    print('insctl_pwr_vset_ch2 = ' + str(insctl_pwr_vset_ch2))
    print('insctl_pwr_iset_ch2 = ' + str(insctl_pwr_iset_ch2))
    print('insctl_pwr_outs_ch2 = ' + str(insctl_pwr_outs_ch2))
    print('insctl_pwr_vset_ch3 = ' + str(insctl_pwr_vset_ch3))
    print('insctl_pwr_iset_ch3 = ' + str(insctl_pwr_iset_ch3))
    print('insctl_pwr_outs_ch3 = ' + str(insctl_pwr_outs_ch3))

    pass


def status_update_pwr():
    # update the status to the related excel table
    sh_inst_ctrl.range((inssts_pwr_connection_y,
                       inssts_pwr_connection_x)).value = pwr_connection_status
    sh_inst_ctrl.range(
        (inssts_pwr_refresh_y, inssts_pwr_refresh_x)).value = insctl_pwr_refresh
    sh_inst_ctrl.range(
        (inssts_pwr_serial_y, inssts_pwr_serial_x)).value = insctl_pwr_serial
    sh_inst_ctrl.range((inssts_pwr_calibration_y,
                       inssts_pwr_calibration_x)).value = insctl_pwr_calibration
    sh_inst_ctrl.range(
        (inssts_pwr_vset_ch1_y, inssts_pwr_vset_ch1_x)).value = pwr_v_ch1_status
    sh_inst_ctrl.range(
        (inssts_pwr_iset_ch1_y, inssts_pwr_iset_ch1_x)).value = pwr_i_ch1_status
    sh_inst_ctrl.range(
        (inssts_pwr_outs_ch1_y, inssts_pwr_outs_ch1_x)).value = pwr_o_ch1_status
    sh_inst_ctrl.range(
        (inssts_pwr_auto_ch1_y, inssts_pwr_auto_ch1_x)).value = pwr_auto_ch1_status

    sh_inst_ctrl.range(
        (inssts_pwr_vset_ch2_y, inssts_pwr_vset_ch2_x)).value = pwr_v_ch2_status
    sh_inst_ctrl.range(
        (inssts_pwr_iset_ch2_y, inssts_pwr_iset_ch2_x)).value = pwr_i_ch2_status
    sh_inst_ctrl.range(
        (inssts_pwr_outs_ch2_y, inssts_pwr_outs_ch2_x)).value = pwr_o_ch2_status
    sh_inst_ctrl.range(
        (inssts_pwr_auto_ch2_y, inssts_pwr_auto_ch2_x)).value = pwr_auto_ch2_status

    sh_inst_ctrl.range(
        (inssts_pwr_vset_ch3_y, inssts_pwr_vset_ch3_x)).value = pwr_v_ch3_status
    sh_inst_ctrl.range(
        (inssts_pwr_iset_ch3_y, inssts_pwr_iset_ch3_x)).value = pwr_i_ch3_status
    sh_inst_ctrl.range(
        (inssts_pwr_outs_ch3_y, inssts_pwr_outs_ch3_x)).value = pwr_o_ch3_status
    sh_inst_ctrl.range(
        (inssts_pwr_auto_ch3_y, inssts_pwr_auto_ch3_x)).value = pwr_auto_ch3_status

    pass


def para_update_load():

    global insctl_load_refresh
    global insctl_load_iset_ch1
    global insctl_load_mset_ch1
    global insctl_load_outs_ch1
    global insctl_load_iset_ch2
    global insctl_load_mset_ch2
    global insctl_load_outs_ch2
    global insctl_load_iset_ch3
    global insctl_load_mset_ch3
    global insctl_load_outs_ch3
    global insctl_load_iset_ch4
    global insctl_load_mset_ch4
    global insctl_load_outs_ch4

    # DC loader table
    insctl_load_refresh = sh_inst_ctrl.range(
        (int(insref_load_y) + 1, int(insref_load_x) + 1)).value
    insctl_load_iset_ch1 = sh_inst_ctrl.range(
        (int(insref_load_y) + 4, int(insref_load_x) + 0)).value
    insctl_load_mset_ch1 = sh_inst_ctrl.range(
        (int(insref_load_y) + 6, int(insref_load_x) + 0)).value
    insctl_load_outs_ch1 = sh_inst_ctrl.range(
        (int(insref_load_y) + 8, int(insref_load_x) + 0)).value
    insctl_load_iset_ch2 = sh_inst_ctrl.range(
        (int(insref_load_y) + 4, int(insref_load_x) + 2)).value
    insctl_load_mset_ch2 = sh_inst_ctrl.range(
        (int(insref_load_y) + 6, int(insref_load_x) + 2)).value
    insctl_load_outs_ch2 = sh_inst_ctrl.range(
        (int(insref_load_y) + 8, int(insref_load_x) + 2)).value
    insctl_load_iset_ch3 = sh_inst_ctrl.range(
        (int(insref_load_y) + 4, int(insref_load_x) + 4)).value
    insctl_load_mset_ch3 = sh_inst_ctrl.range(
        (int(insref_load_y) + 6, int(insref_load_x) + 4)).value
    insctl_load_outs_ch3 = sh_inst_ctrl.range(
        (int(insref_load_y) + 8, int(insref_load_x) + 4)).value
    insctl_load_iset_ch4 = sh_inst_ctrl.range(
        (int(insref_load_y) + 4, int(insref_load_x) + 6)).value
    insctl_load_mset_ch4 = sh_inst_ctrl.range(
        (int(insref_load_y) + 6, int(insref_load_x) + 6)).value
    insctl_load_outs_ch4 = sh_inst_ctrl.range(
        (int(insref_load_y) + 8, int(insref_load_x) + 6)).value

    print('para_update_load')
    print('insctl_load_refresh = ' + str(insctl_load_refresh))
    print('insctl_load_iset_ch1 = ' + str(insctl_load_iset_ch1))
    print('insctl_load_mset_ch1 = ' + str(insctl_load_mset_ch1))
    print('insctl_load_outs_ch1 = ' + str(insctl_load_outs_ch1))
    print('insctl_load_iset_ch2 = ' + str(insctl_load_iset_ch2))
    print('insctl_load_mset_ch2 = ' + str(insctl_load_mset_ch2))
    print('insctl_load_outs_ch2 = ' + str(insctl_load_outs_ch2))
    print('insctl_load_iset_ch3 = ' + str(insctl_load_iset_ch3))
    print('insctl_load_mset_ch3 = ' + str(insctl_load_mset_ch3))
    print('insctl_load_outs_ch3 = ' + str(insctl_load_outs_ch3))
    print('insctl_load_iset_ch4 = ' + str(insctl_load_iset_ch4))
    print('insctl_load_mset_ch4 = ' + str(insctl_load_mset_ch4))
    print('insctl_load_outs_ch4 = ' + str(insctl_load_outs_ch4))

    pass


def status_update_load():

    # update the status to the related excel table
    sh_inst_ctrl.range((inssts_load_connection_y,
                       inssts_load_connection_x)).value = load_connection_status
    sh_inst_ctrl.range(
        (inssts_load_refresh_y, inssts_load_refresh_x)).value = insctl_load_refresh

    # ch1
    sh_inst_ctrl.range(
        (inssts_load_iset_ch1_y, inssts_load_iset_ch1_x)).value = load_i_ch1_status
    sh_inst_ctrl.range((inssts_load_mset_ch1_y,
                       inssts_load_mset_ch1_x)).value = load_m_ch1_status
    sh_inst_ctrl.range(
        (inssts_load_outs_ch1_y, inssts_load_outs_ch1_x)).value = load_o_ch1_status
    sh_inst_ctrl.range(
        (inssts_load_auto_ch1_y, inssts_load_auto_ch1_x)).value = load_auto_ch1_status

    # ch2
    sh_inst_ctrl.range(
        (inssts_load_iset_ch2_y, inssts_load_iset_ch2_x)).value = load_i_ch2_status
    sh_inst_ctrl.range((inssts_load_mset_ch2_y,
                       inssts_load_mset_ch2_x)).value = load_m_ch2_status
    sh_inst_ctrl.range(
        (inssts_load_outs_ch2_y, inssts_load_outs_ch2_x)).value = load_o_ch2_status
    sh_inst_ctrl.range(
        (inssts_load_auto_ch2_y, inssts_load_auto_ch2_x)).value = load_auto_ch2_status

    # ch3
    sh_inst_ctrl.range(
        (inssts_load_iset_ch3_y, inssts_load_iset_ch3_x)).value = load_i_ch3_status
    sh_inst_ctrl.range((inssts_load_mset_ch3_y,
                       inssts_load_mset_ch3_x)).value = load_m_ch3_status
    sh_inst_ctrl.range(
        (inssts_load_outs_ch3_y, inssts_load_outs_ch3_x)).value = load_o_ch3_status
    sh_inst_ctrl.range(
        (inssts_load_auto_ch3_y, inssts_load_auto_ch3_x)).value = load_auto_ch3_status

    # ch4
    sh_inst_ctrl.range(
        (inssts_load_iset_ch4_y, inssts_load_iset_ch4_x)).value = load_i_ch4_status
    sh_inst_ctrl.range((inssts_load_mset_ch4_y,
                       inssts_load_mset_ch4_x)).value = load_m_ch4_status
    sh_inst_ctrl.range(
        (inssts_load_outs_ch4_y, inssts_load_outs_ch4_x)).value = load_o_ch4_status
    sh_inst_ctrl.range(
        (inssts_load_auto_ch4_y, inssts_load_auto_ch4_x)).value = load_auto_ch4_status

    pass


def para_update_met1():

    global insctl_met1_refresh
    global insctl_met1_mset
    global insctl_met1_leve

    insctl_met1_refresh = sh_inst_ctrl.range(
        (int(insref_met1_y) + 1, int(insref_met1_x) + 1)).value
    insctl_met1_mset = sh_inst_ctrl.range(
        (int(insref_met1_y) + 3, int(insref_met1_x) + 0)).value
    # the measurement level of the meter
    insctl_met1_leve = sh_inst_ctrl.range(
        (int(insref_met1_y) + 5, int(insref_met1_x) + 0)).value

    print('para_update_met1')
    print('insctl_met1_refresh = ' + str(insctl_met1_refresh))
    print('insctl_met1_mset = ' + str(insctl_met1_mset))
    print('insctl_met1_leve = ' + str(insctl_met1_leve))

    pass


def status_update_met1():

    global met1_mode_status
    global met1_level_status
    global met1_v_mea_status
    global met1_i_mea_status

    # if insctl_met1_mset == 0:
    #     # enter the voltage measurement mode

    #     met1_i_mea_status = 'NA'
    #     met1_mode_status = 'votlage'
    #     met1_level_status = insctl_met1_leve
    #     pass

    # elif insctl_met1_mset == 1:
    #     # enter the current measurement mode

    #     met1_v_mea_status = 'NA'
    #     met1_mode_status = 'current'
    #     met1_level_status = insctl_met1_leve
    #     pass

    # 220521 some of the status variable comes from the main program, so need other variable save for the result
    sh_inst_ctrl.range(
        (inssts_met1_connection_y, inssts_met1_connection_x)).value = met1_connection_status
    sh_inst_ctrl.range((inssts_met1_refresh_y,
                       inssts_met1_refresh_x)).value = insctl_met1_refresh
    sh_inst_ctrl.range(
        (inssts_met1_mset_y, inssts_met1_mset_x)).value = met1_mode_status
    sh_inst_ctrl.range(
        (inssts_met1_leve_y, inssts_met1_leve_x)).value = met1_level_status
    sh_inst_ctrl.range(
        (inssts_met1_meav_y, inssts_met1_meav_x)).value = met1_v_mea_status
    sh_inst_ctrl.range(
        (inssts_met1_meai_y, inssts_met1_meai_x)).value = met1_i_mea_status

    pass


def para_update_met2():

    global insctl_met2_refresh
    global insctl_met2_mset
    global insctl_met2_leve

    insctl_met2_refresh = sh_inst_ctrl.range(
        (int(insref_met2_y) + 1, int(insref_met2_x) + 1)).value
    insctl_met2_mset = sh_inst_ctrl.range(
        (int(insref_met2_y) + 3, int(insref_met2_x) + 0)).value
    # the measurement level of the meter
    insctl_met2_leve = sh_inst_ctrl.range(
        (int(insref_met2_y) + 5, int(insref_met2_x) + 0)).value

    print('para_update_met2')
    print('insctl_met2_refresh = ' + str(insctl_met2_refresh))
    print('insctl_met2_mset = ' + str(insctl_met2_mset))
    print('insctl_met2_leve = ' + str(insctl_met2_leve))

    pass


def status_update_met2():

    global met2_mode_status
    global met2_level_status
    global met2_v_mea_status
    global met2_i_mea_status

    # if insctl_met2_mset == 0:
    #     # enter the voltage measurement mode

    #     met2_i_mea_status = 'NA'
    #     met2_mode_status = 'votlage'
    #     met2_level_status = insctl_met2_leve
    #     pass

    # elif insctl_met2_mset == 1:
    #     # enter the current measurement mode

    #     met2_v_mea_status = 'NA'
    #     met2_mode_status = 'current'
    #     met2_level_status = insctl_met2_leve
    #     pass

    sh_inst_ctrl.range(
        (inssts_met2_connection_y, inssts_met2_connection_x)).value = met2_connection_status
    sh_inst_ctrl.range((inssts_met2_refresh_y,
                       inssts_met2_refresh_x)).value = insctl_met2_refresh
    sh_inst_ctrl.range(
        (inssts_met2_mset_y, inssts_met2_mset_x)).value = met2_mode_status
    sh_inst_ctrl.range(
        (inssts_met2_leve_y, inssts_met2_leve_x)).value = met2_level_status
    sh_inst_ctrl.range(
        (inssts_met2_meav_y, inssts_met2_meav_x)).value = met2_v_mea_status
    sh_inst_ctrl.range(
        (inssts_met2_meai_y, inssts_met2_meai_x)).value = met2_i_mea_status

    pass


def para_update_src():

    global insctl_src_refresh
    global insctl_src_cset
    global insctl_src_mset
    global insctl_src_leve
    global insctl_src_outs

    insctl_src_refresh = sh_inst_ctrl.range(
        (int(insref_src_y) + 1, int(insref_src_x) + 1)).value
    # setting of the clamp parameter
    insctl_src_cset = sh_inst_ctrl.range(
        (int(insref_src_y) + 3, int(insref_src_x) + 0)).value
    insctl_src_mset = sh_inst_ctrl.range(
        (int(insref_src_y) + 5, int(insref_src_x) + 0)).value
    # the setting level of the source meter
    insctl_src_leve = sh_inst_ctrl.range(
        (int(insref_src_y) + 7, int(insref_src_x) + 0)).value
    # source meter ona and off control status
    insctl_src_outs = sh_inst_ctrl.range(
        (int(insref_src_y) + 13, int(insref_src_x) + 0)).value

    print('para_update_src')
    print('insctl_src_refresh = ' + str(insctl_src_refresh))
    print('insctl_src_cset = ' + str(insctl_src_cset))
    print('insctl_src_mset = ' + str(insctl_src_mset))
    print('insctl_src_leve = ' + str(insctl_src_leve))
    print('insctl_src_outs = ' + str(insctl_src_outs))

    pass


def status_update_src():

    sh_inst_ctrl.range(
        (inssts_src_connection_y, inssts_src_connection_x)).value = src_connection_status
    sh_inst_ctrl.range(
        (inssts_src_refresh_y, inssts_src_refresh_x)).value = insctl_src_refresh
    sh_inst_ctrl.range(
        (inssts_src_cset_y, inssts_src_cset_x)).value = src_clamp_status
    sh_inst_ctrl.range(
        (inssts_src_mset_y, inssts_src_mset_x)).value = src_mode_status
    sh_inst_ctrl.range(
        (inssts_src_leve_y, inssts_src_leve_x)).value = src_level_status
    sh_inst_ctrl.range(
        (inssts_src_meav_y, inssts_src_meav_x)).value = src_v_mea_status
    sh_inst_ctrl.range(
        (inssts_src_meai_y, inssts_src_meai_x)).value = src_i_mea_status
    sh_inst_ctrl.range(
        (inssts_src_outs_y, inssts_src_outs_x)).value = src_o_status

    pass


def check_refresh():
    # this sub is used prevent the dead loop of latch refresh setting

    global insctl_pwr_refresh
    global insctl_load_refresh
    global insctl_met1_refresh
    global insctl_met2_refresh
    global insctl_src_refresh

    insctl_pwr_refresh = sh_inst_ctrl.range(
        (int(insref_pwr_y) + 1, int(insref_pwr_x) + 1)).value
    insctl_load_refresh = sh_inst_ctrl.range(
        (int(insref_load_y) + 1, int(insref_load_x) + 1)).value
    insctl_met1_refresh = sh_inst_ctrl.range(
        (int(insref_met1_y) + 1, int(insref_met1_x) + 1)).value
    insctl_met2_refresh = sh_inst_ctrl.range(
        (int(insref_met2_y) + 1, int(insref_met2_x) + 1)).value
    insctl_src_refresh = sh_inst_ctrl.range(
        (int(insref_src_y) + 1, int(insref_src_x) + 1)).value

    # also need to update the refresh status to the excel table
    # so people know if refresh command is enter or not

    sh_inst_ctrl.range(
        (inssts_pwr_refresh_y, inssts_pwr_refresh_x)).value = insctl_pwr_refresh

    sh_inst_ctrl.range(
        (inssts_load_refresh_y, inssts_load_refresh_x)).value = insctl_load_refresh

    sh_inst_ctrl.range((inssts_met1_refresh_y,
                       inssts_met1_refresh_x)).value = insctl_met1_refresh

    sh_inst_ctrl.range((inssts_met2_refresh_y,
                       inssts_met2_refresh_x)).value = insctl_met2_refresh

    sh_inst_ctrl.range(
        (inssts_src_refresh_y, inssts_src_refresh_x)).value = insctl_src_refresh

    print('check_refresh')
    print('insctl_pwr_refresh = ' + str(insctl_pwr_refresh))
    print('insctl_load_refresh = ' + str(insctl_load_refresh))
    print('insctl_met1_refresh = ' + str(insctl_met1_refresh))
    print('insctl_met2_refresh = ' + str(insctl_met2_refresh))
    print('insctl_src_refresh = ' + str(insctl_src_refresh))

    pass


def inst_name_sheet(nick_name, full_name):
    # definition of sub program may not need the self, but definition of class will need the self
    # self is usually used for internal parameter of class
    # this function will get the nick name and full name from main and update to the sheet
    # based on the nick name

    if nick_name == 'PWR1':
        sh_main.range('D27').value = full_name

    elif nick_name == 'MET1':
        sh_main.range('D28').value = full_name

    elif nick_name == 'MET2':
        sh_main.range('D29').value = full_name

    elif nick_name == 'LOAD1':
        sh_main.range('D30').value = full_name

    elif nick_name == 'LOADSR':
        sh_main.range('D31').value = full_name

    elif nick_name == 'chamber':
        sh_main.range('D32').value = full_name

# this sub used to input SWIRE counter and return related ideal V

# ========== the subprogram been used for more then one file needed
# once this function is called, the new file built with adding extra name
# to generate the new file using this control sheet,
# need to call this function with extra name, so new file will create and mapped to current operating
# main program

# ===
# after the function is called, all the workbooks and worksheets will be assigned
# ===


def build_file(extra_name):
    # build new file and re-load parameters
    global wb_res
    global result_book_trace
    # global sh_main
    # 220824: sh_main is not global here, sh_main will be back to
    # eff_inst after this subprogram is finished
    # to recover, just return the global command
    global sh_org_tab
    global sh_org_tab2
    global sh_org_tab3
    # reserve the sh_inst_ctrl, since still keep the control setting after the auto testing
    # is finished, so control of instrument is from the wb but not wb_res
    # remember to keep wb open always
    # global sh_inst_ctrl
    global wb
    global sheet_arry
    global excel_temp
    global sub_sh_count
    global full_result_name
    # global extra_name

    # define new result workbook
    wb_res = xw.Book()
    # create reference sheet (for sheet position)
    sh_ref = wb_res.sheets.add('ref_sh')
    # delete the extra sheet from new workbook, difference from version
    wb_res.sheets('工作表1').delete()

    # define the sheets in control book
    result_sheet_name = 'raw_out'
    sh_main = wb.sheets('main')
    sh_org_tab = wb.sheets('V_I_com')
    sh_org_tab2 = wb.sheets(result_sheet_name)
    sh_org_tab3 = wb.sheets('I2C_ctrl')
    # copy the sheets to new book
    sh_main.copy(sh_ref)
    sh_org_tab.copy(sh_ref)
    sh_org_tab2.copy(sh_ref)
    sh_org_tab3.copy(sh_ref)

    # assign both sheet to the new sheets in result book
    sh_main = wb_res.sheets('main')
    print(sh_main)
    sh_org_tab = wb_res.sheets('V_I_com')
    sh_org_tab2 = wb_res.sheets(result_sheet_name)
    sh_org_tab3 = wb_res.sheets('I2C_ctrl')
    # sh_main = sh_main.copy(sh_ref)
    # sh_org_tab = sh_org_tab.copy(sh_ref)
    # delete reference after copied

    # 20220112: here is only loading the sheet with command
    # the result initialization will be update after the parameter
    # are loaded into python
    # the result only copy at the end of she control

    # old settings for delete the reference sheet here
    # 20220112 => moe to the end of sheet control because it may add many sheet
    # based on the conrtrol setting loaded from the excel
    # sh_ref.delete()

    # update the result book trace
    full_result_name = new_file_name + '_' + str(extra_name)
    result_book_trace = excel_trace + full_result_name + '.xlsx'

    # save the result book and turn off the control book
    wb_res.save(result_book_trace)
    # wb.close()
    # close control books

    # base on output format copied from the control book
    # start parameter initialization

    # used array to setup the result sheet
    # res_sheet_array = np.zeros(100)
    # this method not working, skip this time

    # here is to generate the sheet
    # counter selection for sheet generation: EL power only one cycle needed (IAVDD= 0)
    if channel_mode == 0 or channel_mode == 1:
        # both only EL or only AVDD just one time
        c_sheet_copy = 1
        if channel_mode == 1:
            sub_sh_count = 4
            # eff + raw + AVDD regulation + Vout regulation
            # 220825 add vout sheet
        elif channel_mode == 0:
            sub_sh_count = 6
            # eff + raw + ELVDD regulation + ELVSS regulation + Vout regulation + Von regulation
    elif channel_mode == 2:
        c_sheet_copy = c_avdd_load
        sub_sh_count = 6
        # eff + raw + ELVDD regulation + ELVSS regulation + Vout regulation + Von regulation

    x_sheet_copy = 0
    sh_temp = sh_org_tab2
    # this loop build the extra sheet needed in the program
    # there are one raw data sheet and fixed format summarize table
    # need to build both efficiency and load regulation summarize table in this loop
    # 3-channel efficiency
    while x_sheet_copy < c_sheet_copy:
        # issue: if x_sheet_copy == 0:

        if channel_mode == 2 or channel_mode == 0:
            # sheet needed for 3 chand only EL are the same
            # just define the different sheet name

            # load AVDD current parameter
            excel_temp = str(sh_org_tab.range(3 + x_sheet_copy, 4).value)

            # =======
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving
            if channel_mode == 2:
                # 3-ch operation
                sheet_temp = 'EFF_I_AVDD=' + excel_temp + 'A'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                    3 + x_sheet_copy, 4).value
            else:
                # EL operation
                sheet_temp = 'EFF'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = '0'
                # no AVDD current, but channel turn on in this operation
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy] = sheet_temp
            sh_org_tab2.name = sheet_temp

            # =======

            # =======
            if channel_mode == 2:
                # add another sheet for the raw data of each AVDD current
                sheet_temp = 'RAW_I_AVDD=' + excel_temp + 'A'
            else:
                sheet_temp = 'RAW'
            # raw data sheet no need example format, can use empty sheet
            sheet_arry[sub_sh_count * x_sheet_copy + 1] = sheet_temp
            wb_res.sheets.add(sheet_temp)
            # =======
            # 220825 explanation added: since the sheet of raw data doesn't have specific
            # format and input needed, add the sheet directly, no need to copy
            # this is the reason why it's different with other sheet generation
            # to add the Vout and Von load regulation, use the format in excel raw_out
            # and it's general format for the regulation and plot function in VBA

            # =======
            # add another sheet for the ELVDD data of each AVDD current
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving
            if channel_mode == 2:
                # 3-ch operation
                sheet_temp = 'ELVDD_I_AVDD=' + excel_temp + 'A'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                    3 + x_sheet_copy, 4).value
            else:
                # EL operation
                sheet_temp = 'ELVDD'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = '0'
                # no AVDD current, but channel turn on in this operation
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy + 2] = sheet_temp
            sh_org_tab2.name = sheet_temp
            # =======

            # =======
            # add another sheet for the ELVSS data of each AVDD current
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving
            if channel_mode == 2:
                # 3-ch operation
                sheet_temp = 'ELVSS_I_AVDD=' + excel_temp + 'A'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                    3 + x_sheet_copy, 4).value
            else:
                # EL operation
                sheet_temp = 'ELVSS'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = '0'
                # no AVDD current, but channel turn on in this operation
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy + 3] = sheet_temp
            sh_org_tab2.name = sheet_temp
            # =======

            # =======
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving
            if channel_mode == 2:
                # 3-ch operation
                sheet_temp = 'Vop_I_AVDD=' + excel_temp + 'A'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                    3 + x_sheet_copy, 4).value
            else:
                # EL operation
                sheet_temp = 'Vop'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = '0'
                # no AVDD current, but channel turn on in this operation
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy + 4] = sheet_temp
            sh_org_tab2.name = sheet_temp

            # =======

            # =======
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving
            if channel_mode == 2:
                # 3-ch operation
                sheet_temp = 'Von_I_AVDD=' + excel_temp + 'A'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = sh_org_tab.range(
                    3 + x_sheet_copy, 4).value
            else:
                # EL operation
                sheet_temp = 'Von'
                # assign the AVDD settting to blue blank of the sheet
                sh_org_tab2.range(21, 3).value = '0'
                # no AVDD current, but channel turn on in this operation
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy + 5] = sheet_temp
            sh_org_tab2.name = sheet_temp

            # =======

        elif channel_mode == 1:
            # sheet build up for only AVDD

            # =======
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving

            # AVDD operation
            sheet_temp = 'EFF'
            # assign the AVDD settting to blue blank of the sheet
            sh_org_tab2.range(21, 3).value = 'NA'
            # here is for AVDD eff
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy] = sheet_temp
            sh_org_tab2.name = sheet_temp

            # =======

            # =======
            # add another sheet for the raw data of each AVDD current
            sheet_temp = 'RAW'
            # raw data sheet no need example format, can use empty sheet
            sheet_arry[sub_sh_count * x_sheet_copy + 1] = sheet_temp
            wb_res.sheets.add(sheet_temp)
            # =======

            # =======
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving

            # AVDD operation
            sheet_temp = 'AVDD'
            # assign the AVDD settting to blue blank of the sheet
            sh_org_tab2.range(21, 3).value = 'NA'
            # here is for AVDD eff
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy + 2] = sheet_temp
            sh_org_tab2.name = sheet_temp

            # =======

            # =======
            sh_temp.copy(sh_ref)
            sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
            # here is to open a new sheet for data saving

            # AVDD operation
            sheet_temp = 'Vout'
            # assign the AVDD settting to blue blank of the sheet
            sh_org_tab2.range(21, 3).value = 'NA'
            # here is for AVDD eff
            # save the sheet name into the array for loading
            sheet_arry[sub_sh_count * x_sheet_copy + 3] = sheet_temp
            sh_org_tab2.name = sheet_temp

            # =======

        x_sheet_copy = x_sheet_copy + 1

    sh_temp.delete()
    # don't need the original raw output format, remove the output

    # 20221012 new sheet delete place is here
    sh_ref.delete()

    wb_res.save(result_book_trace)
    # sub program needed for the contol book initialization


def ideal_v_table(c_swire):
    ideal_v_res = sh_org_tab.range((11 + c_swire, 3)).value
    return ideal_v_res
# this sub is to look for ideal Vin setting for the efficiency testing
# just littlebit faster XD

# add the efficiency re-run sub-prog


def eff_rerun():
    global eff_done_sh
    global sh_org_tab
    global sh_org_tab2
    global sh_org_tab3
    global sh_inst_ctrl
    # this program check the status of the excel file eff_re-run block
    # and update the eff_done to restart efficienct testing
    # from the main, this sub will run if eff_done is already 1
    eff_reset_temp = sh_main.range('B13').value
    print('wait for re-run, update command and setup then set re-run to 1')
    print('the program will start again')

    if eff_reset_temp == 1:
        eff_done_sh = 0
        # reset to 0 if eff sheet is ready to re-run
        # also need to set te input blank back to 0
        sh_main.range('B13').value = 0
        # other wise there will be infinite loop

        # also need to re-assign the mapping sheet to Eff_inst
        # the sheet assignment is gone after finished one round
        re_assign_sheet()

        pass
    else:
        # no need for the action of changing the reset status
        pass

    pass


def re_assign_sheet():
    # this program is to re-assign sheet to prevent loading the data from previous sheet,
    # all the setting and parameter should comes from the Eff_inst

    global sh_org_tab
    global sh_org_tab2
    global sh_org_tab3
    global sh_inst_ctrl

    wb = xw.books('Eff_inst.xlsm')
    result_sheet_name = 'raw_out'
    # sh_main = wb.sheets('main')
    # sheet main is already assign and keep for Eff_inst => main
    sh_org_tab = wb.sheets('V_I_com')
    sh_org_tab2 = wb.sheets(result_sheet_name)
    sh_org_tab3 = wb.sheets('I2C_ctrl')
    sh_inst_ctrl = wb.sheets('inst_ctrl')

    pass


def fully_auto_start():
    # this sub is going to choose to skip all he pop up
    # window or follow the single settings
    # if set to 1 bypass all the pop out function
    global en_plot_waring
    global en_start_up_check
    global en_plot_waring

    # prepare for the
    if en_fully_auto == 1:
        en_plot_waring = 0
        en_start_up_check = 0

    pass
    print('fully auto mode enable')


def fully_auto_end():
    # function TBD, not sure if needed or not
    pass


# below part is the testing for this py file, only operating when this py
# is used for main program
if __name__ == '__main__':

    if test_mode == 0:

        # first is to pring all the parameter and check if the content format is correct
        print('excel related parameter')
        print(control_book_trace)
        print(result_book_trace)
        print(new_file_name)
        print(wb)
        # print(wb_res)
        print(sh_main)
        print(sh_org_tab)
        print('')
        print('pre- test settings ')
        print(pre_test_en)
        print(pre_vin)
        print(pre_vin_max)
        print(pre_imax)
        print('')
        print('GPIB address list')
        print(pwr_sup_addr)
        print(met_vdd_addr)
        print(met_vss_addr)
        print(loa_dts_addr)
        print('')
        print('optional control parameter')
        print(v_clamp_load)
        print(met_vin_rang)
        print(met_iin_rang)
        print('')
        print('other control parameter')
        print(mcu_com_addr)
        print(wait_time)
        print(channel_mode)
        print(pre_inc_vin)
        # wb.close()
        # wb_res.close()
        print('')
        print('parameter output finished, start for sub program ')
        print('')
        print('send the name parameter to related blank of result excel')
        inst_name_sheet('PWR1', 'test1_1')
        inst_name_sheet('MET1', 'test1_2')
        inst_name_sheet('MET2', 'test1_3')
        inst_name_sheet('LOAD1', 'test1_4')
        # wb_res.save(result_book_trace)
        x = 0
        # loop parameter
        input()
        while x < file_count:
            # if x > 0:
            build_file(str(x))
            # wb_res.close()
            # close the sheet at the main

            sh_main.range('E2').value = str(x)
            # add the index mark for checking if the program generate the file
            # with data able to input
            wb_res.save(result_book_trace)
            wb_res.close()
            x = x + 1
            # check on the add file function when control sheet testing

        print('check the excel result')
        input()

        print('')
        print('')
        print('')

    elif test_mode == 1:
        # the test mode for instrument control

        check_refresh()
        # first updatte the internal parameter from the excel sheet settings
        # and print the refresh setting of each instrument

        # update the parameter of each instrument based on the refresh settings
        para_update_pwr()
        # first is the power supply

        # for the status update testing
        # need fake return value for the related status variable

        pwr_connection_status = 'con_sim_pwr'

        pwr_v_ch1_status = 'v_sim_1'
        pwr_i_ch1_status = 'i_sim_1'
        pwr_o_ch1_status = 'o_sim_1'

        pwr_v_ch2_status = 'v_sim_2'
        pwr_i_ch2_status = 'i_sim_2'
        pwr_o_ch2_status = 'o_sim_2'

        pwr_v_ch3_status = 'v_sim_3'
        pwr_i_ch3_status = 'i_sim_3'
        pwr_o_ch3_status = 'o_sim_3'

        status_update_pwr()
        # usually update the status after the parameter updated to the program memory and
        # set to the instrument
        # real time reading value will be updated in the main program (eff_inst)

        para_update_load()
        # second the load

        load_connection_status = 'con_sim_load'

        load_i_ch1_status = 'i1'
        load_m_ch1_status = 'm1'
        load_o_ch1_status = 'o1'

        load_i_ch2_status = 'i2'
        load_m_ch2_status = 'm2'
        load_o_ch2_status = 'o2'

        load_i_ch3_status = 'i3'
        load_m_ch3_status = 'm3'
        load_o_ch3_status = 'o3'

        load_i_ch4_status = 'i4'
        load_m_ch4_status = 'm4'
        load_o_ch4_status = 'o4'

        status_update_load()

        para_update_met1()
        # third the meter 1
        met1_v_mea_status = 'v1_mea'
        met1_i_mea_status = 'i1_mea'

        status_update_met1()

        para_update_met2()
        # fourth the meter 2
        met2_v_mea_status = 'v2_mea'
        met2_i_mea_status = 'i2_mea'

        status_update_met2()

        para_update_src()
        # final the source meter

        src_mode_status = 'load_src.source_type_o'
        src_clamp_status = 'load_src.clamp_VI_o'
        src_level_status = 'load_src.iset_o'
        src_v_mea_status = "load_src.read('VOLT')"
        src_i_mea_status = "load_src.read('CURR')"
        src_o_status = 'load_src.state_o'

        status_update_src()

    print('end, goodbye')
