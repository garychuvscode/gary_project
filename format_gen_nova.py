# sheet_control
# this file is used to generate the excel format needed for the verification port:

# this is nova version based on NT50374

# generate the related table for saving waveform based on the applied setting or row and column
# also the height and width of the wavefrom can be adjust in the control sheet

from copy import copy
from select import select
import xlwings as xw
# import the excel control package

import numpy as np

excel_temp = 0
sheet_temp = 0
str_temp = 0
position_x_temp = 0
position_y_temp = 0
# temp used variable
sheet_arry = np.full([100], None)
# sheet_arry = np.zeros(100)

# color array
color_default = (0, 200, 0)
color_target = (0, 0, 0)
color_temp = (0, 0, 0)

# ==== for nova version variable ====
y_shift_en = 0
# used to identify the row for waveform or the data_mea, only 0 and 1
shift_y_cell = 18
shift_x_cell = 6
x_shift_count = 0
y_shift_count = 0

# the amount need to skip for waveform cell (refere to the cell table)

# # =================== summary for the porgram flow

# 1. initialization
# 2. color adjustment
# 3. row and column indexing and comments input from the control table
# 4. insert the new row for saving the measured result from scope or instrument
# 5. finished and save the result
# note: nova version is much more complicate compare with general version @ indexing thie table
#       since the combination of the cells are complicate, hard to find good and easy relation
#       for the related cells
# but more range application example is used in this program, check below for more detail

# # =================== summary for the porgram flow

# ======== excel book and sheet operation
# control book loading => new book creation => default parameter loading

control_book_trace = 'c:\\py_gary\\test_excel\\general_waveform_table_nova.xlsm'
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

# nova version don't need the width and heigh adjustment
# target_width = sh_main.range('I2').column_width
# target_heigh = sh_main.range('I2').row_height
# default_width = sh_main.range('J5').column_width
# default_heigh = sh_main.range('J5').row_height


# start to adjust the the format based on the input settings

# before changin the format, adjust the color
y_dim = 0
y_shift_en = 0
y_cell = 0
# y_cell is the index for y
y_shift_count = 0
x_shift_count = 0
# x_toggle is used for the cell adjustment of waveform row
x_cell = 0
# x cell index

# 0 means the cell is for waveform at y-axis
c_y_dim = 2 * c_column_item
# y dimension 3*c_column_item, because not adding c_data_mea yet
# knowing the amount is enough
while y_dim < c_y_dim:
    # need to check every cell in the effective operating range
    # from the sheet setting in CTRL sheet

    x_dim = 0
    c_x_dim = c_row_item + 1
    x_shift_count = 0

    # change to another format, this is wrong...
    # if y_shift_en == 1:
    #     # easier to check the index result in loop when use single variable
    #     position_y_temp = 4 + y_dim + y_shift_count * y_shift_en * shift_y_cell
    # else:
    #     position_y_temp = 4 + y_dim + \
    #         y_shift_count * (y_dim - 1) * shift_y_cell
    if y_dim == 0:
        position_y_temp = 4
    elif y_dim % 2 == 1:
        # if y_dim is even,
        y_cell = y_cell + shift_y_cell
        position_y_temp = 4 + y_cell
    elif y_dim % 2 == 0:
        y_cell = y_cell + 1
        position_y_temp = 4 + y_cell

    x_cell = 0
    while x_dim < c_x_dim:

        # # for the index adjusment, need to -1 for x_dim > 2
        # if x_dim < 2:
        #     position_x_temp = 1 + x_dim + x_shift_count * shift_x_cell
        # else:
        #     position_x_temp = x_dim + x_shift_count * shift_x_cell
        # for the index adjusment, need to -1 for x_dim > 2
        if x_dim > 1:
            x_cell = x_cell + shift_x_cell
            position_x_temp = 2 + x_cell
        else:
            position_x_temp = 1 + x_dim

        color_temp = sh_org_tab.range((position_y_temp, position_x_temp)).color

        if color_temp == color_default:
            sh_org_tab.range((position_y_temp, position_x_temp)
                             ).color = color_target

        x_dim = x_dim + 1

        # add the index adjustment for row items
        if x_dim > 1:
            # control variable initial at the loop, don't reset here
            x_shift_count = x_shift_count + 1

    y_dim = y_dim + 1
    # # change the toggle status for every y_dim change
    # if y_shift_en == 0:
    #     y_shift_en = 1
    #     y_shift_count = y_shift_count + 1
    # elif y_shift_en == 1:
    #     y_shift_en = 0

# before insert the row, add the content into realted row and column
# use the same definition but different action
y_dim = 0
c_y_dim = c_column_item
# y dimension 3*c_column_item, because not adding c_data_mea yet
# knowing the amount is enough

# filter the error of extra_str = none (error handling)
if extra_str == None:
    extra_str = ''
if item_str == None:
    item_str = ''
if row_str == None:
    row_str = ''
if col_str == None:
    col_str = ''

while y_dim < c_y_dim:
    # need to check every cell in the effective operating range
    # from the sheet setting in CTRL sheet
    str_temp = sh_main.range((43 + y_dim, 2)).value
    excel_temp = col_str + '\n' + str(str_temp) + item_str + '\n' + extra_str
    position_x_temp = 1
    position_y_temp = 4 + y_dim * (shift_y_cell + 1)
    # position is shift y cell + row between(default is 1)
    sh_org_tab.range((position_y_temp, position_x_temp)).value = excel_temp

    x_dim = 0
    x_shift_count = 0
    c_x_dim = c_row_item
    x_cell = 0
    while x_dim < c_x_dim:

        str_temp = sh_main.range((43 + x_dim, 1)).value
        excel_temp = row_str + str(str_temp)
        # for the index adjusment, need to -1 for x_dim > 2
        if x_dim > 0:
            x_cell = x_cell + shift_x_cell
            position_x_temp = 2 + x_cell
        else:
            position_x_temp = 2

        sh_org_tab.range((position_y_temp, position_x_temp)).value = excel_temp

        x_dim = x_dim + 1

        # add the index adjustment for row items
        if x_dim > 1:
            # control variable initial at the loop, don't reset here
            x_shift_count = x_shift_count + 1

    y_dim = y_dim + 1


# add the new row to the related position
x_column_item = 0
while x_column_item < c_column_item:
    # each column item need to have related data result row

    x_data_mea = 0
    while x_data_mea < c_data_mea:
        # will need to insert the row and assign the row index at the same time
        # first to insert the related row with new color
        position_y_temp = 22 + (shift_y_cell + c_data_mea) * x_column_item

        if x_data_mea > 0:
            excel_temp = 22 + (shift_y_cell + c_data_mea) * x_column_item
            print(excel_temp)
            excel_temp = str(int(excel_temp))
            print(excel_temp)
            excel_temp = excel_temp + ':' + excel_temp
            print(excel_temp)
            sh_org_tab.range(excel_temp).copy()
            sh_org_tab.range(excel_temp).insert()
            # sh_org_tab.api.Rows(position_y_temp).Insert()
        # then assign the related data name for related row ( in reverse order )
        excel_temp = sh_main.range((43 + c_data_mea - x_data_mea - 1, 3)).value
        sh_org_tab.range((position_y_temp, 1)).value = excel_temp
        # # keep the added row in the default high, not change due to insert
        # sh_org_tab.range(
        #     (6 + (2 + c_data_mea) * x_column_item, 1)).row_height = default_heigh

        x_data_mea = x_data_mea + 1
        # testing of insert row into related position
        # sh_org_tab.api.Rows(6).Insert()

    x_column_item = x_column_item + 1

# another loop is for color correction of the measurement result row
# due to there are other cells keep in the default color, so the
# new added row is having default color but not target color

x_color_correction = 0
c_color_correction = c_column_item * (shift_y_cell + c_data_mea)
while x_color_correction < c_color_correction:
    position_x_temp = 1
    position_y_temp = 4 + x_color_correction
    sh_org_tab.range((position_y_temp, position_x_temp)
                     ).color = color_target

    x_color_correction = x_color_correction + 1


# sh_org_tab.range('22:22').copy()
# sh_org_tab.range('22:22').insert()


wb_res.save(result_book_trace)
