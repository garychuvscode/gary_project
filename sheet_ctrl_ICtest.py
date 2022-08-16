# sheet_control for IQ measurement
# this file is used to initializtion the control book and gen the result book

# need to check if the import function can load the sheet variable directly to the main py file


import xlwings as xw
# import the excel control package

excel_temp = 0
# temp used variable

# ======== excel book and sheet operation
# control book loading => new book creation => default parameter loading

control_book_trace = 'c:\\py_gary\\test_excel\\IC_test_ctrl.xlsm'
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
sh_main = wb.sheets('Main')
sh_org_tab = wb.sheets('IQ_measured')
sh_ra_tab = wb.sheets('RA_output')
# copy the sheets to new book
sh_main.copy(sh_ref)
sh_org_tab.copy(sh_ref)
sh_ra_tab.copy(sh_ref)

# assign both sheet to the new sheets in result book
sh_main = wb_res.sheets('Main')
sh_org_tab = wb_res.sheets('IQ_measured')
sh_ra_tab = wb_res.sheets('RA_output')
# sh_main = sh_main.copy(sh_ref)
# sh_org_tab = sh_org_tab.copy(sh_ref)
# delete reference after copied
sh_ref.delete()

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
# for the meter in IC_test, vdd used as voltage meter and vss used as current meter
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


# other control parameter
# COM port parameter input
mcu_com_addr = sh_main.range('C50').value
# general delay time
wait_time = sh_main.range('C51').value
# control mode for MCU
mcu_mode = sh_main.range('C52').value
# pre-increase for efficiency measurement
pre_inc_vin = sh_main.range('C53').value
# vin set for IQ measurement (no need, pre-vin is enough and it will change)
# vin1_set = sh_org_tab.range('C54').value
iq_range = sh_main.range('C54').value
loa_en = sh_main.range('C55').value
loa_curr = sh_main.range('C56').value
loa_curr2 = sh_main.range('C57').value
tab_copy_en = sh_main.range('C58').value

# parameter for the copy table loading
tab_cp_x_sta = sh_main.range('F3').value
tab_cp_y_sta = sh_main.range('G3').value
tab_cp_amount = sh_main.range('H3').value
tab_cp_sheet = sh_main.range('E5').value
tab_cp_x_ran = sh_main.range('F5').value
tab_cp_y_ran = sh_main.range('G5').value
tab_cp_dist = sh_main.range('C59').value

# for RA related test
total_board = sh_main.range('C60').value
pulse_ELVDD = sh_main.range('C61').value
pulse_ELVSS = sh_main.range('C62').value
pulse_AVDD = sh_main.range('C63').value
before_after = sh_main.range('C64').value
Vin = sh_main.range('C65').value
single_test = sh_main.range('C66').value
ISD_hand = sh_main.range('C67').value

# parameter needed from the result sheet (other voltage and current settings)

iq_scaling = sh_org_tab.range('C5').value

# counteer is usually use c_ in opening
c_iq = sh_org_tab.range('C4').value


# RA sheet parameter loading
# none control variable is needed from RA output, but need to put
# the busy flag and polling the number based on the input result from the sheet
# this will be put in the main program but not the sheet control


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

# this sub used to generated the table need to create copy for related sheet
def table_copy(start_x, start_y, amount0, sheet_name, x_size, y_size, dist):
    x_amount = 0
    while x_amount < (amount0 - 1):
        # because there are already one table

        # start to build the copy of table
        y_sta_new = start_y + (y_size + dist) * (x_amount + 1)
        # start point only shift y-axis but not x-axis
        x_sta_new = start_x

        x_x_size = 0
        # x axis update first
        while x_x_size < x_size:

            x_y_size = 0
            # then the y axis
            while x_y_size < y_size:
                if x_y_size == 0 and x_x_size == 0:
                    sh_org_tab.range(
                        y_sta_new - 1, x_sta_new).value = x_amount + 2

                # new element value is copy from the old
                sh_org_tab.range((y_sta_new + x_y_size, x_sta_new + x_x_size)
                                 ).value = sh_org_tab.range((start_y + x_y_size, start_x + x_x_size)).value

                x_y_size = x_y_size + 1
                # adjust counter at the end of loop

            x_x_size = x_x_size + 1
            # adjust counter at the end of loop

        x_amount = x_amount + 1
        # adjust counter at the end of loop

        wb_res.save()
        # save the workbook after copy of table is finished


# the program neede for the the sheet initialization
if tab_copy_en == 1 and single_test != 3:
    # need to start the table copy function
    table_copy(tab_cp_x_sta, tab_cp_y_sta, tab_cp_amount,
               tab_cp_sheet, tab_cp_x_ran, tab_cp_y_ran, tab_cp_dist)
    print('table copy finished')
    # input()


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

    print(iq_range)
    print(loa_en)
    print(loa_curr)
    print(loa_curr2)
    print(tab_copy_en)

    print('check second sheet parameter ')
    print(c_iq)
    print(iq_scaling)

    print('end, goodbye')
    wb_res.save(result_book_trace)
