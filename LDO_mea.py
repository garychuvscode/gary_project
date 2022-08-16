# LDO measurement
# this program used to measure the line and load regulation of LDO
# to improve LDO testing procedure


# general import package

from multiprocessing.connection import wait
from webbrowser import Elinks
from win32con import MB_SYSTEMMODAL
# for the jump out window
import win32api
# also for the jump out window, same group with win32con
import time
# control item related with clock: EX: delay
import pyvisa
# soul of GPIB and MCU control through with COM port
import xlwings as xw
# soul of excel related control, record, parameter loading
import sheet_ctrl_LDO_mea as sh
# loading of information of contorol book and parameter
# import inst_pkg
# loading for instrument object ()
# controlled by simulation mode variable

import locale as lo
# include for atof function => transfer string to float

wait_time = sh.wait_time
wait_small = 0.2
# general waiting time for slppe function

v_res_temp = 0
# measurement result temp variable
pulse1 = 0
pulse2 = 0

# program status temp string
pro_status_str = ''

sim_real = 0
# 1 => real mode, 0=> simulation(program debug) mode
sim_mcu = 1
# this is only used for MCU com port test simuation
# when setting (sim_real, sim_mcu) = (0, 1)
# enable MCU testing when GPIC disable
# because NCU will be separate with GPIB for implementation and test

mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
# MCU array mapping definition for different IO control for relay
meter_ch_ctrl = 0
# meter channel indicator: 0: Vin, 1: Vbias, 2: Vout

mcu_mode_swire = 1
mcu_mode_sw_en = 3
mcu_mode_I2C = 4
mcu_mode_8_bit_IO = 5
mcu_mode_pat_gen_py = 6
mcu_mode_pat_gen_encode = 7
mcu_mode_pat_gen_direct = 8
# MCU mapping for different mode control in 2553

# 20220103 added
vin_ch = 0
# vin channel mapping, default 0 sine calibration will be the first one
v_target = 0
# the vin target for power supply
vbias_target = 0
# target of VBIAs setting
iload_target = 0
# target of loading current

# active sheet (for the result and raw)
sheet_active = ''
raw_active = ''
# the excel table gap for the data in raw sheet
raw_gap = 4 + 5
# setting of raw gap is the " gap + element "
sim_v_data_temp = 0


def mcu_write(index):
    # the MCU write used to generate the command string and send the command out
    # but the related control parameter need to define in the main program
    # here is only used to reduce the code of generate the string
    # and the MCU UART sending command

    if index == 'swire':
        uart_cmd_str = (chr(1) + chr(pulse1) + chr(pulse2))
        # for the SWIRE mode of 2553, there are 2 pulse send to the MCU and DUT
        # pulse amount is from 1 to 255, not sure if 0 will have error or not yet
        # 20220121
    elif index == 'en_sw':
        uart_cmd_str = (chr(3) + chr(pulse1) + chr(1))
        # for the EN SWIRE control mode, need to handle the recover to normal mode (EN, SW) = (1, 1)
        # at the end of application
        # this mode only care about the first data ( 0-4 )
    elif index == 'relay':
        uart_cmd_str = (chr(5) + mcu_cmd_arry[meter_ch_ctrl])

    # print the command going to send before write to MCU, used for debug
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)
    # give some response time for the UART command send and MCU action
    time.sleep(wait_small)


# definition of the output data saveing subprogram

# used for the loading the data to related excel sheet and blank
def data_latch(data_name, mea_res):
    # raw_active.range(( 11 + raw_gap * x_vin, 2 )).value = 'Vin'
    # raw_active.range(( 12 + raw_gap * x_vin, 2 )).value = 'Iin'
    # raw_active.range(( 13 + raw_gap * x_vin, 2 )).value = 'Vout'
    # raw_active.range(( 14 + raw_gap * x_vin, 2 )).value = 'Iout'

    if data_name == 'vin':
        raw_active.range((11 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)

    elif data_name == 'iin':
        raw_active.range((12 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)

    elif data_name == 'vout':
        raw_active.range((13 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        sheet_active.range((25 + x_iload, 3 + x_vin)).value = lo.atof(mea_res)

    elif data_name == 'iout':
        raw_active.range((14 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)

    elif data_name == 'vbias':
        raw_active.range((15 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)


# definition of the program statue update to the excel main sheet
def program_status(status_sting):
    # transfer to the string for following operation
    status_sting_sub = str(status_sting)
    sh.sh_main.range((3, 2)).value = status_sting_sub
    print('status_update: ' + status_sting_sub)
    # use for debugging for the program status update
    # input()

# definition of the calibration sub program


def vin_clibrate_singal_met(vin_ch, vin_target):
    global v_res_temp
    vin_diff = 0
    v_res_temp_f = 0
    # 20220103 add the Vin adjustment function
    # change the measure channel to the Vin channel (No.3)
    # this part added after the load turned on for better Vin setting
    mcu_com.write(chr(5) + mcu_cmd_arry[vin_ch])
    print(chr(5) + mcu_cmd_arry[vin_ch])
    time.sleep(wait_time)
    # measure the first Vin after relay change
    v_res_temp = met1.mea_v()
    v_res_temp_f = lo.atof(v_res_temp)
    # the Vin calibration starts from here

    vin_diff = vin_target - v_res_temp_f
    vin_new = vin_target
    while vin_diff > sh.vin_diff_set or vin_diff < (-1 * sh.vin_diff_set):
        vin_new = vin_new + 0.5 * (vin_target - v_res_temp_f)
        # clamp for the Vin maximum
        if vin_new > sh.pre_vin_max:
            vin_new = sh.pre_vin_max
        pwr1.change_V(vin_new)
        time.sleep(wait_time)
        # measure the Vin change result
        v_res_temp = met1.mea_v()
        v_res_temp_f = lo.atof(v_res_temp)
        vin_diff = vin_target - v_res_temp_f

    # after the loop is finished, record the Vin meausred result (for full load) at the end
    # sh.sh_org_tab.range((10, 9)).value = lo.atof(v_res_temp)
    # the Vin calibration ends from here

    # after the Vin calibration is finished, change the measure channel back
    mcu_com.write(chr(5) + mcu_cmd_arry[meter_ch_ctrl])
    print(chr(5) + mcu_cmd_arry[meter_ch_ctrl])
    time.sleep(wait_time)
    # finished getting back to the initial state


# initialization for the instrument ====

# give the parameter not regular change at definition
# define all the instrument first, then initial settings


rm = pyvisa.ResourceManager()
# not in sim because MCU also need to used

# COM port initial settings
uart_cmd_str = 0
# UART command string for double checking
if sim_mcu == 1:
    uart_cmd_str = "COM" + str(int(sh.mcu_com_addr))
    print(uart_cmd_str)
    mcu_com = rm.open_resource(uart_cmd_str)
    # watch out the form of the address using for connect COM port or GPIB

    # write example
    # mcu_com.write('abcde')
    # read example
    # mcu_res = mcu_com.read()

    # MCU initial for AVDD measurement
    mcu_com.write(chr(mcu_mode_8_bit_IO) + mcu_cmd_arry[meter_ch_ctrl])
    print(chr(mcu_mode_8_bit_IO) + mcu_cmd_arry[meter_ch_ctrl])
    time.sleep(wait_time)
    # input()

# MCU need to cover the change from positive to negative
# ELVDD, ELVSS and AVDD channel change
# MCU channel, initial setting:


# real mode using instrument connection
if sim_real == 1:
    import inst_pkg as inst

    pwr1 = inst.LPS_505N(0, 0, 1, sh.pwr_sup_addr, 'off')
    # no worries about the channel of supply, will change at initial state below
    # def __init__(self, vset0, iset0, act_ch0, GP_addr0, state0):

    met1 = inst.Met_34460(sh.met_vin_res, sh.met_vin_rang,
                          sh.met_iin_res, sh.met_iin_rang, sh.met_vdd_addr)
    # meter loaded from the main sheet
    # def __init__(self, mea_v_res0, max_mea_v0, mea_i_res0, max_mea_i0, GP_addr0):

    load1 = inst.chroma_63600(
        1, sh.loa_dts_addr, str(sh.loa_mod_set))
    # load channel and related setting also change after definition
    # def __init__(self, act_ch0, GP_addr0, mode0):

    # turn on the GPIB connection of instrument
    pwr1.open_inst()
    met1.open_instr()
    load1.open_inst()

    # # load the name to related blank in result book
    # sh.sh_main.range('D27').value = pwr1.inst_name()
    # sh.sh_main.range('D28').value = met1.inst_name()
    # sh.sh_main.range('D30').value = load1.inst_name()

    # the other way for update instrument name(sub program in sh_ctrl)
    sh.inst_name_sheet('PWR1', pwr1.inst_name())
    sh.inst_name_sheet('MET1', met1.inst_name())
    sh.inst_name_sheet('LOAD1', load1.inst_name())

    # settings need for related sheet of instrument

    # protection and other setting (before channel on)

    # power supply OV and OC protection
    pwr1.ov_oc_set(sh.pre_vin_max, sh.pre_imax)

    # power supply channel (channel on setting)
    if sh.pre_test_en == 1:
        pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.pwr_ch_set, 'on')
        # print('pre-power on here')
        if (sh.vbias_en == 1):
            pwr1.chg_out(sh.vbias_pre, sh.vbias_curr_limit,
                         sh.pwr_ch2_set, 'on')
        print('pre-power on here')

        msg_res = win32api.MessageBox(
            0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

    # the power will change from initial state directly, not turn off between transition
    if (sh.vbias_en == 1):
        pwr1.chg_out(sh.vbias_pre, sh.vbias_curr_limit, sh.pwr_ch2_set, 'on')

    pwr1.chg_out(sh.vin_pre, sh.vin_curr_limit, sh.pwr_ch_set, 'on')

    # loader channel and current
    # default off, will be turn on and off based on the loop control

    time.sleep(wait_time)
    # turn the Vin and vbias on, then turn the load on
    load1.chg_out(sh.load_pre, sh.loa_ch_set, 'on')
    # load set for LDO

    # call vin calibration
    vin_clibrate_singal_met(vin_ch, sh.pre_vin)
    # after vin calibration, the v_res_temp will be the last vin value
    # assign to related record for Vin
    sh.sh_org_tab.range((3, 9)).value = lo.atof(v_res_temp)
    time.sleep(wait_time)

    # assign the VBIAS measurement to record
    v_res_temp = met1.mea_v()
    time.sleep(wait_time)
    sh.sh_org_tab.range((2, 9)).value = lo.atof(v_res_temp)

    # turn off the load after measurement finished
    load1.chg_out(sh.load_pre, sh.loa_ch_set, 'off')
    program_status('real test mode done')


else:
    # update the pre-test double check window, see if the pre-test works
    if sh.pre_test_en == 1:
        # pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.pwr_ch_set, 'on')
        print('pre-power on here')

        msg_res = win32api.MessageBox(0, 'press enter if hardware configuration is correct',
                                      'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

    print('pre-power on state finished and ready for next')

    # sh.sh_main.range((3, 2)).value = 'test mode power on ok'
    program_status('test mode power on ok')

    # sh.sh_org_tab.range((10, 6)).value = '3.305'
    # input()

# program measurement loop

x_bias = 0
# counter for SWIRE pulse amount
while x_bias < sh.c_bias:

    if (sh.vbias_en == 1):

        vbias_target = sh.sh_org_tab.range((3 + x_bias, 4)).value
        pro_status_str = 'VBIAS: ' + str(vbias_target)
        program_status(pro_status_str)
        # relay channel selection or initialization of inner loop

        # after loading the Vbias target, change the Vbias voltage settings

        # assign the sheet setting for the result and the raw sheet
        sheet_active = sh.wb_res.sheets(sh.sheet_arry[2 * x_bias])
        raw_active = sh.wb_res.sheets(sh.sheet_arry[2 * x_bias + 1])

        if sim_real == 1:
            # adjust the VBIAS voltage
            pwr1.chg_out(vbias_target, sh.vbias_curr_limit,
                         sh.pwr_ch2_set, 'on')
            time.sleep(wait_small)
        else:
            # simulation mode change bias settings
            print('bias setting change: ' + str(vbias_target))

    else:
        sheet_active = sh.sh_org_tab2
        raw_active = sh.wb_res.sheets.add('raw')

    # Vin loop start point
    # counter for the Vin setting

    x_vin = 0
    while x_vin < sh.c_vin:

        v_target = sh.sh_org_tab.range((3 + x_vin, 2)).value
        pro_status_str = 'Vin:' + str(v_target)
        program_status(pro_status_str)
        # update the target Vin to the program status

        # add the related Vin setting at the result sheet
        sheet_active.range((24, 3 + x_vin)).value = v_target
        # for the raw data sheet index
        raw_active.range((11 + raw_gap * x_vin, 2)).value = 'Vin'
        raw_active.range((12 + raw_gap * x_vin, 2)).value = 'Iin'
        raw_active.range((13 + raw_gap * x_vin, 2)).value = 'Vout'
        raw_active.range((14 + raw_gap * x_vin, 2)).value = 'Iout'
        raw_active.range((15 + raw_gap * x_vin, 2)).value = 'Vbias'

        if sim_real == 1:
            # adjust the vin voltage
            pwr1.chg_out(v_target, sh.vin_curr_limit, sh.pwr_ch_set, 'on')
            time.sleep(wait_small)
        else:
            # simulation mode change bias settings
            print('vin setting change: ' + str(v_target))

        # current loop start point
        # counter of current setting
        x_iload = 0
        while x_iload < sh.c_iload:

            iload_target = sh.sh_org_tab.range((3 + x_iload, 3)).value
            pro_status_str = 'VBIAS: ' + \
                str(vbias_target) + ' vin: ' + str(v_target) + \
                ' i_load: ' + str(iload_target)
            program_status(pro_status_str)
            if x_vin == 0:
                sheet_active.range((25 + x_iload, 2)).value = iload_target
                # only process once for the current index at the result sheet

            if sim_mcu == 1:
                # setting up for the relay channel

                # uart_cmd_str = (chr(mcu_mode_8_bit_IO) +
                #                 mcu_cmd_arry[int(meter_ch_ctrl)])
                # print(uart_cmd_str)
                # mcu_com.write(uart_cmd_str)
                meter_ch_ctrl = 0
                mcu_write('relay')
                time.sleep(wait_small)
                # input()
                # finished adjust relay channel

                # note that the meter_ch_ctrl will change through the measurement channel change
                # so the the start point of meter should be Vin(calibration)
            else:
                print('change the relay channel without MCU for calibration')
                # setting up for the relay channel
                uart_cmd_str = (chr(mcu_mode_8_bit_IO) +
                                mcu_cmd_arry[int(meter_ch_ctrl)])
                print(uart_cmd_str)
                # mcu_com.write(uart_cmd_str)
                time.sleep(wait_small)

            if sim_real == 1:
                # setting up for the instrument control
                time.sleep(wait_small)

                # turn the load on to setting i_load
                load1.chg_out(iload_target, sh.loa_ch_set, 'on')

                # need to set muc_sim to 1 before using calibration
                vin_clibrate_singal_met(vin_ch, v_target)
                # record Vin after calibration finished
                time.sleep(wait_time)
                data_latch('vin', v_res_temp)

                meter_ch_ctrl = meter_ch_ctrl + 1
                # change relay to Vbias
                mcu_write('relay')

                v_res_temp = met1.mea_v()
                time.sleep(wait_time)
                data_latch('vbias', v_res_temp)

                meter_ch_ctrl = meter_ch_ctrl + 1
                mcu_write('relay')
                # change relay to Vout
                v_res_temp = met1.mea_v()
                time.sleep(wait_time)
                data_latch('vout', v_res_temp)

                v_res_temp = pwr1.read_iout()
                # for current reading, need to remove A in the end of string
                v_res_temp = v_res_temp.replace('A', '')
                # this part can also consider to move to the next ints_pkg file
                # can help to improve the complexity
                data_latch('iin', v_res_temp)
                time.sleep(wait_time)
                v_res_temp = load1.read_iout(sh.loa_ch_set)
                # for current reading, need to remove A in the end of string
                v_res_temp = v_res_temp.replace('A', '')
                data_latch('iout', v_res_temp)
                time.sleep(wait_time)
                # record iin and iout for from the source and load
                # meter channel shift, when the cycle end, reset the channel
                # selection for the next round
                meter_ch_ctrl = 0
                mcu_write('relay')
                # after setting the meter channel, also give MCU command back to initial
                # get ready for the next cycle

                # measure Vout
                # measure Iin (power supply or meter)
                # mrasure I out (loader or meter)

                # release loading
                load1.chg_out(iload_target, sh.loa_ch_set, 'off')
                # prepare for the next round

                # save the result after each counter finished
            else:
                # comments for the simulation mode
                # by using different digit of the floating number to know
                # the mapping of variable and settings
                # checking if the result saving is correct or not
                v_res_temp = v_res_temp + 1
                sim_v_data_temp = v_res_temp + v_target
                # add one variable call sim_v_data is because not to effect the v_res counter
                # during saving different result mapping and
                # and also know the program still going forward

                # MCU testing command
                if sim_mcu == 1:
                    # testing for the MCU operation independently
                    meter_ch_ctrl = meter_ch_ctrl + 1
                    # change relay to Vbias
                    mcu_write('relay')
                    print('change relay, meter channel change to: ' +
                          str(meter_ch_ctrl))
                    print('now is vbias')
                    # input()

                    meter_ch_ctrl = meter_ch_ctrl + 1
                    mcu_write('relay')
                    # change relay to Vout
                    print('change relay, meter channel change to: ' +
                          str(meter_ch_ctrl))
                    print('now is vout')
                    # input()

                    meter_ch_ctrl = 0
                    mcu_write('relay')
                    print('change relay, meter channel change to: ' +
                          str(meter_ch_ctrl))
                    print('now is back to initial')
                    # input()

                data_latch('vin', str(sim_v_data_temp))
                v_res_temp = v_res_temp + 1
                sim_v_data_temp = v_res_temp + 0.01
                data_latch('iin', str(sim_v_data_temp))
                v_res_temp = v_res_temp + 1
                sim_v_data_temp = v_res_temp + 0.02
                data_latch('vout', str(sim_v_data_temp))
                v_res_temp = v_res_temp + 1
                sim_v_data_temp = v_res_temp + iload_target
                data_latch('iout', str(sim_v_data_temp))
                v_res_temp = v_res_temp + 1
                sim_v_data_temp = v_res_temp + 0.03
                data_latch('vbias', str(sim_v_data_temp))

                # the simulation mode for result generation
            x_iload = x_iload + 1

        x_vin = x_vin + 1

    sh.wb_res.save(sh.result_book_trace)

    x_bias = x_bias + 1

# turn off the load and source after the loop is finished
if (sh.vbias_en == 1):
    pwr1.chg_out(0, sh.vbias_curr_limit, sh.pwr_ch2_set, 'off')
pwr1.chg_out(0, sh.vin_curr_limit, sh.pwr_ch_set, 'off')
load1.chg_out(0, sh.loa_ch_set, 'off')

print('finsihed and goodbye')
