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
# only the main program use this method to separate the excel and program
# since here will include the loop structure and overall control of report generation
import sheet_ctrl_main_obj as sh
import parameter_load_obj as par

# ==============
# excel setting for main program

# define the excel object
excel_obj = par.excel_parameter(str(sh.file_setting))
excel_obj.open_result_book()


# excel setting for main program
# ==============
excel_obj.save()
# ==============
# main program structure

# instrument initialization
# all the default had beent fix in the program and change directly in definition below


# # reference sheet delete after all the test finished
# # delete reference after copied
# sh.sh_ref.delete()
# sh.sh_ref_condition.delete()
# sh.wb_res.save()
# 220901 change to delete the reference sheet in excel object
excel_obj.end_of_test()

# main program structure
# ==============
