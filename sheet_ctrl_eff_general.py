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

# ======== excel book and sheet operation
# control book loading => new book creation => default parameter loading

control_book_trace = 'c:\\py_gary\\test_excel\\Eff_general.xlsm'
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


wb = xw.books('Eff_general.xlsm')
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

# # update the result book trace
# result_book_trace = excel_trace + new_file_name + str(extra_name) + '.xlsx'
result_book_trace = ''
full_result_name = ''

# # save the result book and turn off the control book
# # 220324 move to build file
# wb_res.save(result_book_trace)
# # wb.close()
# # close control books

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

# optional control parameter for instrument
v_clamp_load = sh_main.range('C37').value
# this is for source meter
met_vin_rang = sh_main.range('C38').value
# maximum V for meter
met_iin_rang = sh_main.range('C39').value
# maximum I for meter
pwr_ch_set = sh_main.range('C40').value
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
    global sh_main
    global sh_org_tab
    global sh_org_tab2
    global sh_org_tab3
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
            sub_sh_count = 3
            # eff + raw + AVDD regulation
        elif channel_mode == 0:
            sub_sh_count = 4
            # eff + raw + ELVDD regulation + ELVSS regulation
    elif channel_mode == 2:
        c_sheet_copy = c_avdd_load
        sub_sh_count = 4
        # eff + raw + ELVDD regulation + ELVSS regulation

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


# below part is the testing for this py file, only operating when this py
# is used for main program
if __name__ == '__main__':

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
    print('end, goodbye')
