# this is the master program and excel file
# include below control items:
# 1. excel => insturment address, pre-test control patameter
# result file trace and new file name settings
# 2. program => system loop and sequence arragement, instrument initialization,
# result sheet saving (after each function)

import xlwings as xw

import win32com.client
# this import is for the VBA function

# the import function
# ==============


# ==============
# excel operating setting

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
sh_ref = wb_res.sheets.add('ref_sh')
# delete the extra sheet from new workbook, difference from version
wb_res.sheets('工作表1').delete()

# define the sheets in control book

# excel operating setting
# ==============


# ==============
# main program structure

# main program structure
# ==============
