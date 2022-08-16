# sheet_control
# this file is used to generate the excel format needed for the verification port:

# to have easier indexing of the cells, this format is different with the novw format
# used for the auto testing generation in the future, not related with the nova version

# generate the related table for saving waveform based on the applied setting or row and column
# also the height and width of the wavefrom can be adjust in the control sheet


# need to check if the import function can load the sheet variable directly to the main py file


import xlwings as xw
# import the excel control package

import numpy as np
# numpy is used for matrix operation, check more from google

excel_temp = 0
sheet_temp = 0
str_temp = 0
# temp used variable
sheet_arry = np.full([100], None)
# sheet_arry = np.zeros(100)

# color array
color_default = (0, 200, 0)
color_target = (0, 0, 0)
color_temp = (0, 0, 0)
# either loading the color from original sheet or general the new colcor is ok
# RGB setting for the three parameter in the array

# # =================== summary for the porgram flow

# 1. initialization
# 2. color adjustment
# 3. row and column indexing and comments input from the control table
# 4. insert the new row for saving the measured result from scope or instrument
# 5. finished and save the result

# # =================== summary for the porgram flow


# ======== excel book and sheet operation
# control book loading => new book creation => default parameter loading

control_book_trace = 'c:\\py_gary\\test_excel\\general_waveform_table.xlsm'
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
sheet_name1 = 'CTRL'
sheet_name2 = 'table'
# sheet_name3 = 'V_out'
sh_main = wb.sheets(sheet_name1)
sh_org_tab = wb.sheets(sheet_name2)
# sh_org_tab2 = wb.sheets(sheet_name3)
# copy the sheets to new book
sh_main.copy(sh_ref)
sh_org_tab.copy(sh_ref)
# sh_org_tab2.copy(sh_ref)

# assign both sheet to the new sheets in result book
sh_main = wb_res.sheets(sheet_name1)
sh_org_tab = wb_res.sheets(sheet_name2)
# sh_org_tab2 = wb_res.sheets(sheet_name3)
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
sh_ref.delete()
# delete the reference sheet

excel_temp = str(sh_main.range('C36').value)
# load the trace as string
new_file_name = str(sh_main.range('C35').value)

# update the result book trace
result_book_trace = excel_temp + new_file_name + '.xlsx'


# save the result book and turn off the control book
wb_res.save(result_book_trace)
# wb.close()
# close control books

# loading the control values
c_row_item = sh_main.range('C31').value
c_column_item = sh_main.range('C32').value
c_data_mea = sh_main.range('C33').value

item_str = sh_main.range('C28').value
row_str = sh_main.range('C29').value
col_str = sh_main.range('C30').value
extra_str = sh_main.range('C34').value
color_default = sh_main.range('C37').color
color_target = sh_main.range('C38').color

target_width = sh_main.range('I2').column_width
target_heigh = sh_main.range('I2').row_height
default_width = sh_main.range('J5').column_width
default_heigh = sh_main.range('J5').row_height
# default height and width is used to prevent shape change of the table
# target height and width is used to save the waveform capture from scope

# start to adjust the the format based on the input settings

# before changin the format, adjust the color
y_dim = 0
c_y_dim = 3 * c_column_item
# y dimension 3*c_column_item, because not adding c_data_mea yet
# knowing the amount is enough
while y_dim < c_y_dim:
    # need to check every cell in the effective operating range
    # from the sheet setting in CTRL sheet

    x_dim = 0
    c_x_dim = c_row_item + 1
    while x_dim < c_x_dim:
        color_temp = sh_org_tab.range((4 + y_dim, 1 + x_dim)).color

        if color_temp == color_default:
            sh_org_tab.range((4 + y_dim, 1 + x_dim)).color = color_target
            # change the needed place with default color to the target color

        x_dim = x_dim + 1

    y_dim = y_dim + 1

# before insert the row, add the content into realted row and column
# use the same definition but different action
y_dim = 0
c_y_dim = c_column_item
# y dimension 3*c_column_item, because not adding c_data_mea yet
# knowing the amount is enough

# filter the error of extra_str = none (error handling for the no input cells)
if extra_str == None:
    extra_str = ''
if item_str == None:
    item_str = ''
if row_str == None:
    row_str = ''
if col_str == None:
    col_str = ''

# x_dim, y_dim are the dimension counter for modifing the table
while y_dim < c_y_dim:
    # need to check every cell in the effective operating range
    # from the sheet setting in CTRL sheet
    str_temp = sh_main.range((43 + y_dim, 2)).value
    excel_temp = col_str + '\n' + str(str_temp) + item_str + '\n' + extra_str
    sh_org_tab.range((4 + 1 + y_dim * 3, 1)).value = excel_temp
    sh_org_tab.range((4 + 1 + y_dim * 3, 1)).row_height = target_heigh
    # height need to change when modifing the column cells

    x_dim = 0
    c_x_dim = c_row_item
    while x_dim < c_x_dim:

        str_temp = sh_main.range((43 + x_dim, 1)).value
        excel_temp = row_str + str(str_temp)
        sh_org_tab.range((4 + y_dim * 3, 1 + 1 + x_dim)).value = excel_temp
        sh_org_tab.range((4 + y_dim * 3, 1 + 1 + x_dim)
                         ).column_width = target_width
        # width need to change when modifing the row cells

        x_dim = x_dim + 1

    y_dim = y_dim + 1


# add the new row to the related position
x_column_item = 0
while x_column_item < c_column_item:
    # each column item need to have related data result row

    x_data_mea = 0
    while x_data_mea < c_data_mea:
        # will need to insert the row and assign the row index at the same time
        # first to insert the related row with new color
        if x_data_mea > 0:
            sh_org_tab.api.Rows(6 + (2 + c_data_mea) * x_column_item).Insert()

        # then assign the related data name for related row ( in reverse order )
        excel_temp = sh_main.range((43 + c_data_mea - x_data_mea - 1, 3)).value
        sh_org_tab.range(
            (6 + (2 + c_data_mea) * x_column_item, 1)).value = excel_temp
        # keep the added row in the default high, not change due to insert
        sh_org_tab.range(
            (6 + (2 + c_data_mea) * x_column_item, 1)).row_height = default_heigh

        x_data_mea = x_data_mea + 1
        # testing of insert row into related position
        # sh_org_tab.api.Rows(6).Insert()

    x_column_item = x_column_item + 1

wb_res.save(result_book_trace)
# save the result after program is finished
