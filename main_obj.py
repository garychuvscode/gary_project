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
import mcu_obj as mcu
# for the excel sheet control
# only the main program use this method to separate the excel and program
# since here will include the loop structure and overall control of report generation
import sheet_ctrl_main_obj as sh
import parameter_load_obj as para

# import for the verification object
import IQ_scan_obj as iq


# off line test, set to 1 set all the instrument to simulation mode
main_off_line = 1
# this is the variable control file name, single or the multi item
# adjust after the if selection of program_group
multi_item = 0


# ==============
# excel setting for main program

# define the excel object
excel_m = para.excel_parameter(str(sh.file_setting))

# excel_m.open_result_book()

# excel_m.excel_save()
# excel setting for main program
# ==============


# ==============
# instrument startup configuration
pwr_m = inst.LPS_505N(excel_m.pwr_vset, excel_m.pwr_iset,
                      excel_m.pwr_act_ch, excel_m.pwr_supply_addr, excel_m.pwr_ini_state)
met_i_m = inst.Met_34460(excel_m.met_v_res, excel_m.met_v_max,
                         excel_m.met_i_res, excel_m.met_i_max, excel_m.meter2_i_addr)
# default turn the MCU on
mcu_m = mcu.MCU_control(1, excel_m.mcu_com_addr)
# instrument startup configuration
# ==============


# ==============
# definition of sub program needed in main
def sim_mode_all(main_off_line0):
    # set all the stuff inst into simulation mode
    if main_off_line0 == 0:
        # all the stuff set to experiment mode
        pwr_m.sim_inst = 1
        met_i_m.sim_inst = 1
        mcu_m.sim_mcu = 1

        pass
    else:
        # all the stuff set to simulation mode
        pwr_m.sim_inst = 0
        met_i_m.sim_inst = 0
        mcu_m.sim_mcu = 0
        excel_m.sim_mode_delay(0.02, 0.01)
        pass

    pass


# ==============


# ==============
# main program structure
program_group = excel_m.program_group_index

# check simulation mode or experiment mode
# for single inst change, adjut by each object
# this is the general key
sim_mode_all(main_off_line)

excel_m.open_result_book()

excel_m.excel_save()

# different verififcation combination
# decide by the program_group variable
if program_group == 0:
    # here is the single test for IQ
    multi_item = 0
    # single setting of the object need to be 1

    iq_test = iq.iq_scan(excel_m, pwr_m, excel_m.pwr_act_ch, met_i_m, mcu_m, 1)

    # generate(or copy) the needed sheet to the result book
    iq_test.sheet_gen()

    # start the testing
    iq_test.run_verification()

    # remember that this is only call by main, not by  object
    excel_m.end_of_test(0)

    print('end of the IQ object testing program')

    pass

elif program_group == 1:

    pass


# instrument initialization
# all the default had beent fix in the program and change directly in definition below


# # reference sheet delete after all the test finished
# # delete reference after copied
# sh.sh_ref.delete()
# sh.sh_ref_condition.delete()
# sh.wb_res.save()
# 220901 change to delete the reference sheet in excel object
# excel_m.end_of_test(multi_item)

# main program structure
# ==============
