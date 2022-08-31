# this is the master program and excel file
# include below control items:
# 1. excel => insturment address, pre-test control patameter
# result file trace and new file name settings
# 2. program => system loop and sequence arragement, instrument initialization,
# result sheet saving (after each function)

# === add on import
# for the excel related operation
import xlwings as xw
# this import is for the VBA function
import win32com.client


# === other support py import
# for the instrument objects
import inst_pkg_d as inst
# for the excel sheet control
import main_obj_sheet_control as sh


# ==============
# main program structure

# instrument initialization
# all the default had beent fix in the program and change directly in definition below


# reference sheet delete after all the test finished
# delete reference after copied
sh.sh_ref.delete()
sh.sh_ref_condition.delete()

# main program structure
# ==============
