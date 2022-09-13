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
import SWIRE_scan_obj as sw
import EFF_obj as eff

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
met_v_m = inst.Met_34460(excel_m.met_v_res, excel_m.met_v_max,
                         excel_m.met_i_res, excel_m.met_i_max, excel_m.meter1_v_addr)
met_i_m = inst.Met_34460(excel_m.met_v_res, excel_m.met_v_max,
                         excel_m.met_i_res, excel_m.met_i_max, excel_m.meter2_i_addr)
loader_chr_m = inst.chroma_63600(
    excel_m.loader_act_ch, excel_m.loader_addr, excel_m.loader_ini_mode)
src_m = inst.Keth_2440(excel_m.src_vset, excel_m.src_iset, excel_m.loader_src_addr,
                       excel_m.src_ini_state, excel_m.src_ini_type, excel_m.src_clamp_ini)
chamber_m = inst.chamber_su242(excel_m.cham_tset_ini, excel_m.chamber_addr,
                               excel_m.cham_ini_state, excel_m.cham_l_limt, excel_m.cham_h_limt, excel_m.cham_hyst)

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
        met_v_m.sim_inst = 1
        met_i_m.sim_inst = 1
        loader_chr_m.sim_inst = 1
        src_m.sim_inst = 1
        chamber_m.sim_inst = 1
        mcu_m.sim_mcu = 1

        pass
    else:
        # all the stuff set to simulation mode
        pwr_m.sim_inst = 0
        met_v_m.sim_inst = 0
        met_i_m.sim_inst = 0
        loader_chr_m.sim_inst = 0
        src_m.sim_inst = 0
        chamber_m.sim_inst = 0
        mcu_m.sim_mcu = 0
        excel_m.sim_mode_delay(0.02, 0.01)
        pass

    pass


def sim_mode_independent(pwr, met_v, met_i, loader, src, chamber):
    # independent setting for instrument simulation mode
    if pwr == 1:
        pwr_m.sim_inst = 1
        pass
    else:
        pwr_m.sim_inst = 0
        pass
    if met_v == 1:
        met_v_m.sim_inst = 1
        pass
    else:
        met_v_m.sim_inst = 0
        pass
    if met_i == 1:
        met_i_m.sim_inst = 1
        pass
    else:
        met_i_m.sim_inst = 0
        pass
    if loader == 1:
        loader_chr_m.sim_inst = 1
        pass
    else:
        loader_chr_m.sim_inst = 0
        pass
    if src == 1:
        src_m.sim_inst = 1
        pass
    else:
        src_m.sim_inst = 0
        pass
    if chamber == 1:
        chamber_m.sim_inst = 1
        pass
    else:
        chamber_m.sim_inst = 0
        pass

    pass


def open_inst_and_name():
    # this used to turn all the instrument on after program start
    # setup simulation mode help to prevent error

    # open instrument
    pwr_m.open_inst()
    met_v_m.open_inst()
    met_i_m.open_inst()
    loader_chr_m.open_inst()
    src_m.open_inst()
    chamber_m.open_inst()
    mcu_m.com_open()

    # for the instrument in simulation mode, name will be set to simulation mode

    # link the name to the sheet
    excel_m.inst_name_sheet('PWR1', pwr_m.inst_name())
    excel_m.inst_name_sheet('MET1', met_v_m.inst_name())
    excel_m.inst_name_sheet('MET2', met_i_m.inst_name())
    excel_m.inst_name_sheet('LOAD1', loader_chr_m.inst_name())
    excel_m.inst_name_sheet('LOADSR', src_m.inst_name())
    excel_m.inst_name_sheet('chamber', chamber_m.inst_name())

    pass


def change_file_name(new_file_name_str):
    excel_m.wb.sheets('main').range('B8').value = str(new_file_name_str)

    pass

# ==============


# add the supported verification item and create the related object name
iq_test = iq.iq_scan(excel_m, pwr_m, met_i_m, mcu_m)
sw_test = sw.sw_scan(excel_m, pwr_m, met_v_m, loader_chr_m, mcu_m)
eff_test = eff.eff_mea(excel_m, pwr_m, met_v_m,
                       loader_chr_m, mcu_m, src_m, met_i_m, chamber_m)

# ==============
# main program structure
program_group = excel_m.program_group_index

# check simulation mode or experiment mode
# for single inst change, adjut by each object
# this is the general key
sim_mode_all(main_off_line)

# to make sure moudlize, put in to the selection
# excel_m.open_result_book()
# excel_m.excel_save()

# different verififcation combination
# decide by the program_group variable
if program_group == 0:
    excel_m.open_result_book()
    excel_m.excel_save()
    # here is the single test for IQ

    # single setting of the object need to be 1 => no needed single
    multi_item = 0

    if main_off_line == 0:
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 0, 1, 0, 0, 0)
        pass

    # definition of experiment object
    # iq_test = iq.iq_scan(excel_m, pwr_m, met_i_m, mcu_m)

    # generate(or copy) the needed sheet to the result book
    iq_test.sheet_gen()

    # open instrument and add the name
    open_inst_and_name()

    # start the testing
    iq_test.run_verification()

    # remember that this is only call by main, not by  object
    excel_m.end_of_file(multi_item)

    print('end of the IQ object testing program')

    pass

elif program_group == 1:
    excel_m.open_result_book()
    excel_m.excel_save()
    # SWIRE scan single verififcation

    # single setting of the object need to be 1 => no needed single
    multi_item = 0

    if main_off_line == 0:
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 1, 0, 1, 0, 0)
        pass

    # definition of experiment object
    # sw_test = sw.sw_scan(excel_m, pwr_m, met_v_m, loader_chr_m, mcu_m)

    # generate(or copy) the needed sheet to the result book
    sw_test.sheet_gen()

    # open instrument and add the name
    open_inst_and_name()

    # start the testing
    sw_test.run_verification()

    # remember that this is only call by main, not by  object
    excel_m.end_of_file(multi_item)

    print('end of the IQ object testing program')

    pass

elif program_group == 2:
    excel_m.open_result_book()
    excel_m.excel_save()
    # SWIRE + IQ testing

    # single setting of the object need to be 1 => no needed single
    multi_item = 1

    # if not off line testing, setup the the instrument needed independently
    if main_off_line == 0:
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 1, 1, 1, 0, 0)
        pass

    # definition of experiment object
    # iq_test = iq.iq_scan(excel_m, pwr_m, met_i_m, mcu_m)

    # sw_test = sw.sw_scan(excel_m, pwr_m, met_v_m, loader_chr_m, mcu_m)

    # generate(or copy) the needed sheet to the result book
    sw_test.sheet_gen()
    iq_test.sheet_gen()

    # open instrument and add the name
    open_inst_and_name()

    # start the testing
    iq_test.run_verification()
    sw_test.run_verification()

    # remember that this is only call by main, not by  object
    excel_m.end_of_file(multi_item)

    print('end of the IQ object testing program')

    pass

elif program_group == 3:
    excel_m.open_result_book()
    excel_m.excel_save()
    # efficiency testing ( I2C and SWIRE-normal mode )

    # single setting of the object need to be 1 => no needed single
    multi_item = 0

    if main_off_line == 0:
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 1, 1, 1, 1, 0)
        pass

    # definition of experiment object
    # eff_test = eff.eff_mea(excel_m, pwr_m, met_v_m,
    #                        loader_chr_m, mcu_m, src_m, met_i_m, chamber_m)

    # generate(or copy) the needed sheet to the result book
    eff_test.sheet_gen()
    # excel_m.build_file()

    # open instrument and add the name
    open_inst_and_name()

    # start the testing
    eff_test.run_verification()

    # 220907 test for change name
    # excel_m.detail_name = '_detail name added'

    # remember that this is only call by main, not by  object
    # excel_m.end_of_file(multi_item)

    print('end of the EFF object testing program')

    pass

elif program_group == 4:

    # change file name should be with different setting, cancel the plan to
    # make the change file name function
    # change_file_name('test_name')

    # fixed part, open one result book and save the book
    # in temp name
    excel_m.open_result_book()
    excel_m.excel_save()
    # verification items

    # single setting of the object need to be 1 => no needed single
    multi_item = 1

    # if not off line testing, setup the the instrument needed independently
    if main_off_line == 0:
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 1, 1, 1, 1, 0)
        pass

    # open instrument and add the name
    # must open after simulation mode setting(open real or sim)
    open_inst_and_name()
    print('open instrument with real or simulation mode')

    # changeable area
    # ===========

    # generate(or copy) the needed sheet to the result book
    sw_test.sheet_gen()
    iq_test.sheet_gen()
    eff_test.sheet_gen()
    print('finished sheet generation')

    # start the testing
    # iq_test.run_verification()
    print('IQ test finished')
    # sw_test.run_verification()
    print('SW test finished')
    eff_test.run_verification()
    print('efficiency test finished')

    # ===========
    # changeable area

    # remember that this is only call by main, not by  object
    excel_m.end_of_file(multi_item)
    print('end of the program')

    pass

# reference code
elif program_group == 1000:
    # fixed part, open one result book and save the book
    # in temp name
    excel_m.open_result_book()
    excel_m.excel_save()
    # verification items

    # single setting of the object need to be 1 => no needed single
    multi_item = 1

    # if not off line testing, setup the the instrument needed independently
    if main_off_line == 0:
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 1, 1, 1, 1, 0)
        pass

    # open instrument and add the name
    # must open after simulation mode setting(open real or sim)
    open_inst_and_name()
    print('open instrument with real or simulation mode')

    # changeable area
    # ===========

    # generate(or copy) the needed sheet to the result book

    print('finished sheet generation')

    # start the testing
    # run_verification() => should be put in here

    print('finished XX verification')

    # ===========
    # changeable area

    # remember that this is only call by main, not by  object
    excel_m.end_of_file(multi_item)
    print('end of the program')

    pass


# instrument initialization
# all the default had beent fix in the program and change directly in definition below


# # reference sheet delete after all the test finished
# # delete reference after copied
# sh.sh_ref.delete()
# sh.sh_ref_condition.delete()
# sh.wb_res.save()
# 220901 change to delete the reference sheet in excel object
# excel_m.end_of_file(multi_item)

# main program structure
# ==============
