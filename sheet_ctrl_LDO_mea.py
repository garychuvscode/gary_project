# sheet_control
# this file is used to initializtion the control book and gen the result book

# need to check if the import function can load the sheet variable directly to the main py file


import xlwings as xw
# import the excel control package

import numpy as np

excel_temp = 0
sheet_temp = 0
# temp used variable
sheet_arry = np.full([100], None)
# sheet_arry = np.zeros(100)

# ======== excel book and sheet operation
# control book loading => new book creation => default parameter loading

control_book_trace = 'c:\\py_gary\\test_excel\\LDO_mea.xlsm'
# no place to load the trace from excel or program, define by default
result_book_trace = ''
# result trace unable to load yet

new_file_name = ''
# new file name not loaded yet


# open control workbook
wb = xw.Book(control_book_trace)

# define new result workbook
wb_res = xw.Book()
# create reference sheet (for sheet position)
sh_ref = wb_res.sheets.add('ref_sh')
# delete the extra sheet from new workbook, difference from version
wb_res.sheets('工作表1').delete()

# define the sheets in control book
result_sheet_name = 'V_out'
sh_main = wb.sheets('main')
sh_org_tab = wb.sheets('V_I_com')
sh_org_tab2 = wb.sheets(result_sheet_name)
# copy the sheets to new book
sh_main.copy(sh_ref)
sh_org_tab.copy(sh_ref)
sh_org_tab2.copy(sh_ref)

# assign both sheet to the new sheets in result book
sh_main = wb_res.sheets('main')
sh_org_tab = wb_res.sheets('V_I_com')
sh_org_tab2 = wb_res.sheets(result_sheet_name)
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

excel_temp = str(sh_main.range('B9').value)
# load the trace as string
new_file_name = str(sh_main.range('B8').value)

# update the result book trace
result_book_trace = excel_temp + new_file_name + '.xlsx'

# save the result book and turn off the control book
wb_res.save(result_book_trace)
# wb.close()
# close control books


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

# optional control parameter for instrument
v_clamo_load = sh_main.range('C37').value
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
loa_mod_set = sh_main.range('C44').value
loa_ch2_set = sh_main.range('C45').value
pwr_ch2_set = sh_main.range('C46').value


# other control parameter
# COM port parameter input
mcu_com_addr = sh_main.range('C50').value
# general delay time
wait_time = sh_main.range('C51').value
# control mode for MCU
mcu_mode = sh_main.range('C52').value
# pre-increase for efficiency measurement
pre_inc_vin = sh_main.range('C53').value
# the setting for Vin calibration accuracy
vin_diff_set = sh_main.range('C54').value
# the current limit of the VIBAS channel
vbias_curr_limit = sh_main.range('C55').value
vin_curr_limit = sh_main.range('C56').value

# parameter needed from the result sheet
# setting for the VBIAS, VIN and current, loaded the counter
# into the python

vbias_en = sh_org_tab.range('H1').value
vbias_pre = sh_org_tab.range('H2').value
vin_pre = sh_org_tab.range('H3').value
load_pre = sh_org_tab.range('H4').value

# counteer is usually use c_ in opening
c_bias = sh_org_tab.range('D1').value
c_vin = sh_org_tab.range('B1').value
c_iload = sh_org_tab.range('C1').value


# used array to setup the result sheet
# res_sheet_array = np.zeros(100)
# this method not working, skip this time


if vbias_en == 1:

    # here is to generate the sheet for each VBIAS settings
    c_sheet_copy = c_bias
    x_sheet_copy = 0
    sh_temp = sh_org_tab2
    while x_sheet_copy < c_sheet_copy:
        # if x_sheet_copy == 0:
        sh_temp.copy(sh_ref)
        sh_org_tab2 = wb_res.sheets(result_sheet_name + ' (2)')
        excel_temp = str(sh_org_tab.range(3 + x_sheet_copy, 4).value)
        sheet_temp = 'VBIAS=' + excel_temp
        # save the sheet name into the array for loading
        sheet_arry[2 * x_sheet_copy] = sheet_temp
        sh_org_tab2.name = sheet_temp
        # also assign the value to the VBIAS element
        sh_org_tab2.range(21, 3).value = sh_org_tab.range(
            3 + x_sheet_copy, 4).value
        # add another sheet for the raw data of each VBIAS sheets
        sheet_temp = 'raw_VBAIS=' + excel_temp
        sheet_arry[2 * x_sheet_copy + 1] = sheet_temp
        wb_res.sheets.add(sheet_temp)

        x_sheet_copy = x_sheet_copy + 1

    sh_temp.delete()
else:
    sh_org_tab2 = wb_res.sheets(result_sheet_name)
    # even there are no other bias settings, also need to set the c_bias
    # so the loop of measurement can start
    c_bias = 1

# 20221012 new sheet delete place is here
sh_ref.delete()


wb_res.save(result_book_trace)
# sub program needed for the contol book initialization


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

# this sub used to input SWIRE counter and return related ideal V


def ideal_v_table(c_swire):
    ideal_v_res = sh_org_tab.range((11 + c_swire, 3)).value
    return ideal_v_res


# below part is the testing for this py file, only operating when this py
# is used for main program
if __name__ == '__main__':

    # first is to pring all the parameter and check if the content format is correct
    print('excel related parameter')
    print(control_book_trace)
    print(result_book_trace)
    print(new_file_name)
    # print(wb)
    print(wb_res)
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
    print(v_clamo_load)
    print(met_vin_rang)
    print(met_iin_rang)
    print('')
    print('other control parameter')
    print(mcu_com_addr)
    print(wait_time)
    print(mcu_mode)
    print(pre_inc_vin)
    print('')
    print('parameter output finished, start for sub program ')
    print('')
    print('send the name parameter to related blank of result excel')
    inst_name_sheet('PWR1', 'test1_1')
    inst_name_sheet('MET1', 'test1_2')
    inst_name_sheet('MET2', 'test1_3')
    inst_name_sheet('LOAD1', 'test1_4')
    print('check the excel result')
    input()

    print('')
    print('')
    print('')
    print('end, goodbye')
