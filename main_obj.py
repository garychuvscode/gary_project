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

# extra library
import Scope_LE6100A as sco
import Power_BK9141 as bk

# import for the verification object
import IQ_scan_obj as iq
import SWIRE_scan_obj as sw
import EFF_obj as eff
import instrument_scan_obj as ins_scan
import format_gen_obj as form_g
import general_test_obj as gene_t
import ripple_obj as rip


# off line test, set to 1 set all the instrument to simulation mode
main_off_line = 0
single_mode = 0
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
scope_m = sco.Scope_LE6100A(excel0=excel_m, main_off_line0=main_off_line)
pwr_bk_m = bk.Power_BK9141(
    excel0=excel_m, GP_addr0=excel_m.pwr_bk_addr, main_off_line0=main_off_line)


# instrument startup configuration
# ==============


# ==============
# definition of sub program needed in main
def sim_mode_all(main_off_line0):
    # set all the stuff inst into simulation mode
    if main_off_line0 == 0:
        # all the stuff set to experiment mode

        # 221110: leave the setting based on GPIB address and only
        # operat thwne the main_off_line is 1
        # pwr_m.sim_inst = 1
        # met_v_m.sim_inst = 1
        # met_i_m.sim_inst = 1
        # loader_chr_m.sim_inst = 1
        # src_m.sim_inst = 1
        # chamber_m.sim_inst = 1
        # mcu_m.sim_mcu = 1
        # scope_m.sim_inst = 1

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
        scope_m.sim_inst = 0
        pwr_bk_m.sim_inst = 0
        excel_m.sim_mode_delay(0.02, 0.01)
        pass

    pass

# sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
# scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)


def sim_mode_independent(pwr=0, met_v=0, met_i=0, loader=0, src=0, chamber=0, scope=0, bk_pwr=0, main_off_line0=1, single_mode0=0):
    # independent setting for instrument simulation mode
    '''
    if the test mode = 0, default disable all the single simulation mode function\n
    only based on the setting of GPIB address to decide the setting of simulation mode
    for the
    '''
    if main_off_line0 == 0 and single_mode0 == 1:
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
        if scope == 1:
            scope_m.sim_inst = 1
            pass
        else:
            scope_m.sim_inst = 0
            pass
        if bk_pwr == 1:
            pwr_bk_m.sim_inst = 1
            pass
        else:
            pwr_bk_m.sim_inst = 0
            pass

        pass

    pass


def open_inst_and_name():
    # this used to turn all the instrument on after program start
    # setup simulation mode help to prevent error

    if main_off_line == 1:
        sim_mode_all(main_off_line)
        # all instrument turn to simulation mode

    # open instrument
    pwr_m.open_inst()
    met_v_m.open_inst()
    met_i_m.open_inst()
    loader_chr_m.open_inst()
    src_m.open_inst()
    chamber_m.open_inst()
    mcu_m.com_open()
    scope_m.open_inst()
    pwr_bk_m.open_inst()

    # for the instrument in simulation mode, name will be set to simulation mode

    # link the name to the sheet
    excel_m.inst_name_sheet('PWR1', pwr_m.inst_name())
    excel_m.inst_name_sheet('MET1', met_v_m.inst_name())
    excel_m.inst_name_sheet('MET2', met_i_m.inst_name())
    excel_m.inst_name_sheet('LOAD1', loader_chr_m.inst_name())
    excel_m.inst_name_sheet('LOADSR', src_m.inst_name())
    excel_m.inst_name_sheet('chamber', chamber_m.inst_name())
    excel_m.inst_name_sheet('scope', scope_m.inst_name())
    excel_m.inst_name_sheet('bkpwr', pwr_bk_m.inst_name())

    # pending: think about the name of scope, how to input?

    pass


def change_file_name(new_file_name_str):
    excel_m.wb.sheets('main').range('B8').value = str(new_file_name_str)

    pass


def loader_cal_excel():
    loader_chr_m.current_cal_setup(excel_m.loader_cal_offset_ELch, excel_m.loader_cal_offset_VCIch,
                                   0, 0, excel_m.loader_cal_leakage_ELch, excel_m.loader_cal_leakage_VCIch, 0, 0)
    # turn on the calibration mode
    loader_chr_m.cal_mode_en = 1
    pass

# ==============


# add the supported verification item and create the related object name
iq_test = iq.iq_scan(excel_m, pwr_m, met_i_m, mcu_m)
sw_test = sw.sw_scan(excel_m, pwr_m, met_v_m, loader_chr_m, mcu_m)
eff_test = eff.eff_mea(excel_m, pwr_m, met_v_m,
                       loader_chr_m, mcu_m, src_m, met_i_m, chamber_m)
in_scan = ins_scan.instrument_scan(excel_m, pwr_m, met_v_m,
                                   loader_chr_m, mcu_m, src_m, met_i_m, chamber_m)
format_g = form_g.format_gen(excel_m)
general_t = gene_t.general_test(excel_m, pwr_m, met_v_m,
                                loader_chr_m, mcu_m, src_m, met_i_m, chamber_m)
general_t_bk = gene_t.general_test(excel_m, pwr_bk_m, met_v_m,
                                   loader_chr_m, mcu_m, src_m, met_i_m, chamber_m)

if excel_m.pwr_select == 0:
    # set to 0 is to use LPS505
    ripple_t = rip.ripple_test(excel_m, pwr_m, met_v_m,
                               loader_chr_m, mcu_m, src_m, met_i_m, chamber_m, scope_m)
elif excel_m.pwr_select == 1:
    # set to 1 is to use BK9141
    ripple_t = rip.ripple_test(excel_m, pwr_bk_m, met_v_m,
                               loader_chr_m, mcu_m, src_m, met_i_m, chamber_m, scope_m)

# scope cpature setting for waveform related testing item
if main_off_line == 1:
    ripple_t.obj_sim_mode = 0
    general_t.obj_sim_mode = 0
else:
    ripple_t.obj_sim_mode = 1
    general_t.obj_sim_mode = 1

# ==============

if __name__ == '__main__':
    #  the testing code for this file object

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

    # Single IQ
    if program_group == 0:
        '''
        item for exe file operaing, don't change after confirm with experiment
        '''
        excel_m.open_result_book()
        excel_m.excel_save()
        # here is the single test for IQ
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(pwr=1, met_i=1, main_off_line0=main_off_line)
        # open instrument and add the name
        open_inst_and_name()

        # start the testing
        iq_test.run_verification()

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        print('end of the IQ object testing program')
        pass

    # single SWIRE
    elif program_group == 1:
        '''
        item for exe file operaing, don't change after confirm with experiment
        '''
        excel_m.open_result_book()
        excel_m.excel_save()
        # SWIRE scan single verififcation
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(pwr=1, met_v=1, met_i=1,
                             loader=1, main_off_line0=main_off_line)
        # open instrument and add the name
        open_inst_and_name()

        # start the testing
        sw_test.run_verification()

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        print('end of the SWIRE object testing program')
        pass

    # SWIRE + IQ testing
    elif program_group == 2:
        '''
        item for exe file operaing, don't change after confirm with experiment
        '''
        excel_m.open_result_book()
        excel_m.excel_save()
        # SWIRE + IQ testing

        # single setting of the object need to be 1 => no needed single
        multi_item = 1
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(pwr=1, met_v=1, met_i=1,
                             loader=1, main_off_line0=main_off_line)
        # sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1, scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)

        # open instrument and add the name
        open_inst_and_name()

        # start the testing
        iq_test.run_verification()
        sw_test.run_verification()

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        print('end of the IQ and SWIRE object testing program')
        pass

    # single test for efficiency chamber option
    elif program_group == 3:
        '''
        item for exe file operaing, don't change after confirm with experiment
        '''
        excel_m.open_result_book()
        excel_m.excel_save()
        # efficiency testing ( I2C and SWIRE-normal mode )
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1,
                             src=1, chamber=1, main_off_line0=main_off_line)
        # open instrument and add the name
        open_inst_and_name()

        # start the testing
        eff_test.run_verification()
        # issue for using end of file should be solve for efficiency test

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        print('end of the EFF object testing program')

        pass

    # IQ + SWIRE + efficiency (eff default in 1 file)
    elif program_group == 4:
        '''
        item for exe file operaing, don't change after confirm with experiment
        '''
        excel_m.open_result_book()
        # excel_m.excel_save()
        # verification items

        # single setting of the object need to be 1 => no needed single
        multi_item = 1
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        # sim_mode_independent(1, 1, 1, 1, 1, 0, main_off_line0=main_off_line)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1,
                             src=1, chamber=1, main_off_line0=main_off_line)

        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()

        # changeable area
        # ===========

        # 220921 add the current calibration setting for loader
        loader_chr_m.current_cal_setup(
            excel_m.loader_cal_offset_ELch, excel_m.loader_cal_offset_VCIch, 0, 0, 0, 0, 0, 0)

        # start the testing
        iq_test.run_verification()
        print('IQ test finished')
        sw_test.run_verification()
        print('SW test finished')
        # cancel this line if make eff single file available
        # single file = 1 => all in same file, 0 => all in different file
        excel_m.eff_single_file = 1
        eff_test.run_verification()
        print('efficiency test finished')

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        print('end of the program 4')
        pass

    # single verification, independent file
    elif program_group == 4.5:

        # single setting of the object need to be 1 => no needed single
        multi_item = 1
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 1, 1, 1, 0, 0, main_off_line0=main_off_line)

        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()

        excel_m.open_result_book()
        iq_test.run_verification()
        excel_m.end_of_file(0)

        excel_m.open_result_book()
        sw_test.run_verification()
        excel_m.end_of_file(0)

        excel_m.open_result_book()
        eff_test.run_verification()
        excel_m.end_of_file(0)

        excel_m.open_result_book()
        general_t.set_sheet_name('general_1')
        general_t.run_verification()

        general_t.set_sheet_name('general_2')
        general_t.run_verification()
        excel_m.end_of_file(0)

        pass

    # instrument control panel only
    elif program_group == 5:
        '''
        item for exe file operaing, don't change after confirm with experiment
        '''
        # fixed part, open one result book and save the book
        # in temp name
        # excel_m.open_result_book()
        # verification items

        # single setting of the object need to be 1 => no needed single
        multi_item = 0

        # if not off line testing, setup the the instrument needed independently
        # 221110: using GPIB address to decide simulation mode
        # if main_off_line == 0:
        #     # set simulation for the used instrument
        #     # pwr, met_v, met_i, loader, src, chamber
        #     sim_mode_independent(
        #         1, 1, 1, 1, 1, 0, main_off_line0=main_off_line)
        #     pass

        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========
        while excel_m.program_exit == 1:
            in_scan.check_inst_update()
            # the program exit will be check after the check inst update
            # the loop will break automatically after change the program exit

        print('finished XX verification')

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        # excel_m.end_of_file(multi_item)
        print('end of the program')

        pass

    # testing for current calibration (chroma 63600)
    elif program_group == 5.5:
        # testing for the current calibration of the chroma loader

        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber
        sim_mode_independent(1, 1, 1, 1, 0, 0, main_off_line0=main_off_line)

        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        loader_chr_m.current_calibration(met_i_m, pwr_m, 3, 1, 6.6)

        print('finished the loader calibration, check result')
        # give interrupt for the parameter check
        input()
        loader_chr_m.current_calibration(met_i_m, pwr_m, 3, 2, 3.3)

        # give interrupt for the parameter check
        input()

        pass

    # single test for general test
    elif program_group >= 6 and program_group < 7:
        # fixed part, open one result book and save the book
        '''
        6 => new file, cal_vin
        6.1 => old file, cal_vin
        6.2 => new file
        6.3 => old file
        '''
        # in temp name
        if program_group == 6.1 or program_group == 6.3:
            # track previous report and save at the end
            excel_m.open_result_book(keep_last=1)
        else:
            excel_m.open_result_book(keep_last=0)
        # auto save after the book is generate
        excel_m.excel_save()

        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber, main offline
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        temp_str = str(excel_m.single_test_mapped_general)
        print(f'now is single test for {temp_str}')
        general_t.set_sheet_name(temp_str)
        if program_group == 6:
            general_t.run_verification()
        elif program_group == 6.2 or program_group == 6.3:
            # 6.1 is the version without Vin calibration
            general_t.run_verification(vin_cal=0)

        print('finished general_test verification')

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # single test for waveform capture
    elif program_group == 7:
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book()
        # auto save after the book is generate
        excel_m.excel_save()

        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber, main offline
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========
        temp_str = str(excel_m.single_test_mapped_wave)
        print(f'now is single test for {temp_str}')

        format_g.set_sheet_name(temp_str)

        # add the protection of line transient setting pwr as real mode
        if excel_m.sh_format_gen.range('B15').value == 1:
            # set the pwr to simulation mode
            pwr_m.sim_inst = 0
            pwr_bk_m.sim_inst = 0

        ripple_t.run_verification()
        # format_g.table_return()

        print('finished waveform test verification')

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # For the HV buck ripple, line and load transient
    elif program_group == 7.5:
        excel_m.open_result_book()

        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber, main offline
        sim_mode_independent(1, 1, 1, 1, 1, 0, main_off_line0=main_off_line)
        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========
        temp_str = str(excel_m.single_test_mapped_wave)
        print(f'now is single test for {temp_str}')

        format_g.set_sheet_name(temp_str)

        # add the protection of line transient setting pwr as real mode
        if excel_m.sh_format_gen.range('B15').value == 1:
            # set the pwr to simulation mode
            pwr_m.sim_inst = 0
            pwr_bk_m.sim_inst = 0

        if excel_m.pwr_select == 1:
            # this is only for HV buck
            excel_m.relay0_ch = 1
            excel_m.message_box(
                'high V buck setting, parallel output for BK9141\n control shannel is set to CH1', 'waatch out', auto_exception=1)

        ripple_t.run_verification()

        print('finished waveform test verification')

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # For the HV buck ripple, load transeint in 1
    elif program_group == 7.8:
        excel_m.open_result_book()

        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber, main offline
        sim_mode_independent(1, 1, 1, 1, 1, 0, main_off_line0=main_off_line)
        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        if excel_m.pwr_select == 1:
            # this is only for HV buck
            excel_m.relay0_ch = 1
            excel_m.message_box(
                'high V buck setting, parallel output for BK9141\n control shannel is set to CH1', 'waatch out', auto_exception=1)
        # # ripple
        # format_g.set_sheet_name('CTRL_sh_ripple_SY')
        # ripple_t.run_verification()
        # load transient
        format_g.set_sheet_name('CTRL_sh_load_SY')
        ripple_t.run_verification()

        print('finished waveform test verification')

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # testing for the general test object
    elif program_group == 8:
        '''
        this is the testing can be modify in the program
        by using different code and test \n
        function reserve for VScode control interface,
        don't change previous setting once the testing is ok for items for exe file
        '''
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book()
        # auto save after the book is generate
        # excel_m.excel_save()

        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber, main offline
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1,
                             src=1, chamber=1, main_off_line0=main_off_line)
        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # general_t.set_sheet_name('general_1')
        # general_t.run_verification()

        general_t.set_sheet_name('general_2')
        general_t.run_verification()

        print('finished XX verification')

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # this is going to test for ripple with multi setting and multi items
    # IQ, SWIRE scan, efficiency, ripple => fully auto
    # ISD pending, eff is forced to one file
    elif program_group == 9:
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book(keep_last=1)
        # auto save after the book is generate
        excel_m.excel_save()

        # single setting of the object need to be 1 => no needed single
        multi_item = 1
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber, main offline
        sim_mode_independent(1, 1, 1, 1, 1, 0, main_off_line0=main_off_line)
        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # sheet generation is added in the run verification

        # start the testing

        iq_test.run_verification()
        print('IQ test finished')
        sw_test.run_verification()
        print('SW test finished')
        # cancel this line if make eff single file available
        # single file = 1 => all in same file, 0 => all in different file
        excel_m.eff_single_file = 1
        eff_test.run_verification()
        print('efficiency test finished')
        temp_str = str(excel_m.single_test_mapped_wave)
        print(f'now is single test for {temp_str}')
        format_g.set_sheet_name(temp_str)
        ripple_t.run_verification()
        # format_g.table_return()

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program IQ + SW + eff + ripple')

        pass

    # this is for the flexible ripple related verification
    elif program_group == 10:
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book()
        # auto save after the book is generate
        excel_m.excel_save()

        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # if not off line testing, setup the the instrument needed independently
        # set simulation for the used instrument
        # pwr, met_v, met_i, loader, src, chamber, main offline
        sim_mode_independent(1, 1, 1, 1, 1, 0, main_off_line0=main_off_line)
        # open instrument and add the name
        # must open after simulation mode setting(open real or sim)
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # sheet generation is added in the run verification

        # start the testing

        format_g.set_sheet_name('CTRL_sh_ex_ripple')
        ripple_t.run_verification()
        # format_g.table_return()

        # ===========
        # changeable area

        # remember that this is only call by main, not by  object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # research for the power on and off
    elif program_group == 11:
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book()
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        format_g.set_sheet_name('CTRL_sh_seq_EN=SW')
        ripple_t.pwr_seq()
        # format_g.table_return()
        format_g.set_sheet_name('CTRL_sh_seq_EN')
        ripple_t.pwr_seq()
        # format_g.table_return()
        format_g.set_sheet_name('CTRL_sh_seq_SW')
        ripple_t.pwr_seq()
        # format_g.table_return()
        excel_m.extra_file_name = '_pwr_seq_wave'

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # power sequence output no need others
    elif program_group == 11.1:
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book()
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # used for inrush current measurement
    elif program_group == 12:
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book()
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # temp_str = str(excel_m.single_test_mapped_wave)
        # print(f'now is single test for {temp_str}')
        # format_g.set_sheet_name(temp_str)

        # fix the sheet lock to CTRL_sh_inrush
        format_g.set_sheet_name('CTRL_sh_inrush')
        ripple_t.inrush_current()
        # format_g.table_return()

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # this is used for power on and off sequence half auto
    elif program_group == 13:
        # fixed part, open one result book and save the book
        # in temp name
        excel_m.open_result_book()
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # fix the sheet lock to CTRL_sh_seq_EN=SW, CTRL_sh_seq_EN, CTRL_sh_seq_SW
        format_g.set_sheet_name('CTRL_sh_inrush')
        ripple_t.inrush_current()

        format_g.set_sheet_name('CTRL_sh_seq_EN=SW')
        ripple_t.pwr_seq()
        # format_g.table_return()
        format_g.set_sheet_name('CTRL_sh_seq_EN')
        ripple_t.pwr_seq()
        # format_g.table_return()
        format_g.set_sheet_name('CTRL_sh_seq_SW')
        ripple_t.pwr_seq()
        # format_g.table_return()
        excel_m.extra_file_name = '_inrush_pwr_seq'

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # OTP testing for HV buck
    elif program_group >= 14 and program_group < 15:
        # fixed part, open one result book and save the book
        '''
        explanation of different number settings
        14 => new file
        14.1 => old file
        '''
        # in temp name
        if program_group == 14.1:
            # track previous report and save at the end
            excel_m.open_result_book(keep_last=1)
        else:
            excel_m.open_result_book(keep_last=0)
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        general_t.set_sheet_name('gen_OTP')
        general_t.run_verification(ctrl_ind_1=1)

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # for gen-power on off testing (both on and off different
    # sequence with record)
    elif program_group >= 15 and program_group < 16:
        # fixed part, open one result book and save the book
        '''
        explanation of different number settings
        15 => new file
        15.1 => old file
        '''
        # in temp name
        if program_group == 15.1:
            # track previous report and save at the end
            excel_m.open_result_book(keep_last=1)
        else:
            excel_m.open_result_book(keep_last=0)
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        general_t.set_sheet_name('gen_pwr_on_off_35', 0)
        general_t.set_sheet_name('gen_pwr_on_off_35', 1, '_pwr_off')
        general_t.gen_pwr_on_off()

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # for HV buck chamber related testing
    elif program_group >= 16 and program_group < 17:
        # fixed part, open one result book and save the book
        '''
        explanation of different number settings
        16 => new file
        16.1 => old file
        '''
        # in temp name
        if program_group == 16.1:
            # track previous report and save at the end
            excel_m.open_result_book(keep_last=1)
        else:
            excel_m.open_result_book(keep_last=0)
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # first should be the band gap
        general_t.set_sheet_name(
            ctrl_sheet_name0='gen_BK_band_gap', extra_sheet=0, extra_name='_BK')
        general_t.set_sheet_name(
            ctrl_sheet_name0='gen_BK_band_gap', extra_sheet=1, extra_name='_LDO')
        general_t.run_verification(ctrl_ind_1=2)

        # OTP not toggle EN1
        general_t.set_sheet_name(
            'gen_BK_OTP', extra_sheet=0, extra_name='_EN1_keep')
        general_t.run_verification(ctrl_ind_1=0)
        general_t.set_sheet_name(
            'gen_BK_OTP', extra_sheet=0, extra_name='_EN1_toggle')
        general_t.run_verification(ctrl_ind_1=1)

        # high temp power on off
        general_t.set_sheet_name(
            'gen_BK_pwr_on_off_85', extra_sheet=0, extra_name='_on')
        general_t.set_sheet_name(
            'gen_BK_pwr_on_off_85', extra_sheet=1, extra_name='_off')
        general_t.gen_pwr_on_off()

        # low temp power on off
        general_t.set_sheet_name(
            'gen_BK_pwr_on_off_-40', extra_sheet=0, extra_name='_on')
        general_t.set_sheet_name(
            'gen_BK_pwr_on_off_-40', extra_sheet=1, extra_name='_off')
        general_t.gen_pwr_on_off()

        # file name index
        general_t.extra_file_name_setup('_chamber_mix')

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # for HV buck VTH related testing
    elif program_group >= 17 and program_group < 18:
        # fixed part, open one result book and save the book
        '''
        explanation of different number settings
        17 => new file
        17.1 => old file
        '''
        # in temp name
        if program_group == 17.1:
            # track previous report and save at the end
            excel_m.open_result_book(keep_last=1)
        else:
            excel_m.open_result_book(keep_last=0)
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # 221226 change the current setting to 0.1A to preven burn down
        excel_m.gen_pwr_i_set = 0.1

        # EN=Vin/2 testing
        general_t_bk.set_sheet_name(
            ctrl_sheet_name0='gen_BK_ENd2', extra_sheet=0, extra_name='_')
        general_t.pwr_iout_set(iout_r0=0.1,iout_r6=0.1,iout_r7=0.1)
        general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # EN1 testing
        general_t_bk.set_sheet_name(
            ctrl_sheet_name0='gen_BK_EN1', extra_sheet=0, extra_name='_')
        general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # EN2 testing
        general_t_bk.set_sheet_name(
            ctrl_sheet_name0='gen_BK_EN2', extra_sheet=0, extra_name='_')
        general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # UVLO testing
        general_t_bk.set_sheet_name(
            ctrl_sheet_name0='gen_BK_Vin', extra_sheet=0, extra_name='_')
        general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # file name index
        general_t.extra_file_name_setup('_VTH_mix')

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # testin for BK9141 EN/2
    elif program_group == 100:
        # fixed part, open one result book and save the book
        '''
        explanation of different number settings
        17 => new file
        17.1 => old file
        '''
        # in temp name
        if program_group == 100.1:
            # track previous report and save at the end
            excel_m.open_result_book(keep_last=1)
        else:
            excel_m.open_result_book(keep_last=0)
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        # EN=Vin/2 testing
        general_t_bk.set_sheet_name(
            ctrl_sheet_name0='gen_BK_ENd2', extra_sheet=0, extra_name='_')
        general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # # EN1 testing
        # general_t_bk.set_sheet_name(
        #     ctrl_sheet_name0='gen_BK_EN1', extra_sheet=0, extra_name='_')
        # general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # # EN2 testing
        # general_t_bk.set_sheet_name(
        #     ctrl_sheet_name0='gen_BK_EN2', extra_sheet=0, extra_name='_')
        # general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # # UVLO testing
        # general_t_bk.set_sheet_name(
        #     ctrl_sheet_name0='gen_BK_Vin', extra_sheet=0, extra_name='_')
        # general_t_bk.run_verification(ctrl_ind_1=0, vin_cal=0)

        # file name index
        general_t.extra_file_name_setup('_BK_temp_test')

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass

    # reference code
    elif program_group >= 1000 and program_group < 1001:
        # fixed part, open one result book and save the book
        '''
        explanation of different number settings
        1000 => new file, cal_vin
        1000.1 => old file, cal_vin
        1000.2 => new file
        1000.3 => old file
        '''
        # in temp name
        if program_group == 1000.1 or program_group == 1000.3:
            # track previous report and save at the end
            excel_m.open_result_book(keep_last=1)
        else:
            excel_m.open_result_book(keep_last=0)
        # auto save after the book is generate
        excel_m.excel_save()
        # single setting of the object need to be 1 => no needed single
        multi_item = 0
        # setup instruement for test mode, only for debug, no need to change)
        sim_mode_independent(pwr=1, met_v=1, met_i=1, loader=1, src=1, chamber=1,
                             scope=1, bk_pwr=1, main_off_line0=main_off_line, single_mode0=single_mode)
        # open instrument and add the name to result book
        open_inst_and_name()
        print('open instrument with real or simulation mode')

        # changeable area
        # ===========

        temp_str = str(excel_m.single_test_mapped_general)
        print(f'now is single test for {temp_str}')
        general_t.set_sheet_name(temp_str)
        if program_group == 1000:
            general_t.run_verification()
        elif program_group == 1000.2 or program_group == 1000.3:
            # 6.1 is the version without Vin calibration
            general_t.run_verification(vin_cal=0)

        # ===========
        # changeable area

        # remember that this is only call by main, not by object
        excel_m.end_of_file(multi_item)
        # end of file can also be call between each item
        print('end of the program')

        pass
