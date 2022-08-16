# SWIRE_scan:
# this program used to scan the different SWIRE command output for no load and full load

# general import package

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
import sh_ctrl as sh
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


sim_real = 1
# 1 => real mode, 0=> simulation(program debug) mode
sim_mcu = 1
# this is only used for MCU com port test simuation
# when setting (sim_real, sim_mcu) = (0, 1)
# enable MCU testing when GPIC disable
# because NCU will be separate with GPIB for implementation and test
mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']

meter_ch_ctrl = 1
# meter_ch_ctrl = 0
# meter channel indicator: 0: AVDD, 1: ELVDD, 2: ELVSS, 3: Vin(for calibration)
# 20220508 meter channel indicator: 1: AVDD, 2 ELVDD, 3: ELVSS, 0: Vin(for calibration)

mode = 2
# MCU mapping for RA_GPIO control is mode 2 (SWIRE scan)

# 20220103 added
# vin_ch = 3
vin_ch = 0

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
    mcu_com.write(chr(5) + mcu_cmd_arry[meter_ch_ctrl])
    print(chr(5) + mcu_cmd_arry[meter_ch_ctrl])
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

    met1 = inst.Met_34460(sh.met_vin_res, sh.met_vin_rang,
                          sh.met_iin_res, sh.met_iin_rang, sh.met_vdd_addr)
    # meter loaded from the main sheet

    load1 = inst.chroma_63600(
        1, sh.loa_dts_addr, str(sh.loa_mod_set))
    # load channel and related setting also change after definition

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
        print('pre-power on here')

        msg_res = win32api.MessageBox(
            0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

    # the power will change from initial state directly, not turn off between transition
    pwr1.chg_out(sh.vin1_set, sh.pre_sup_iout, sh.pwr_ch_set, 'on')

    # loader channel and current
    # default off, will be turn on and off based on the loop control

    load1.chg_out(sh.iload1_set, sh.loa_ch_set, 'off')
    # load set for EL-power
    load1.chg_out(sh.iload2_set, sh.loa_ch2_set, 'off')
    # load set for AVDD

    time.sleep(wait_time)
    # AVDD measurement will be independant case, not in the loop
    # keep the loop simple and periodic

    # == AVDD measurement

    # call vin calibration
    vin_clibrate_singal_met(vin_ch, sh.pre_vin)
    # after vin calibration, the v_res_temp will be the last vin value
    # assign to related record
    sh.sh_org_tab.range((10, 9)).value = lo.atof(v_res_temp)

    v_res_temp = met1.mea_v()
    time.sleep(wait_time)
    sh.sh_org_tab.range((10, 4)).value = lo.atof(v_res_temp)

    load1.chg_out(sh.iload2_set, sh.loa_ch2_set, 'on')
    time.sleep(wait_time)

    # call vin calibration
    vin_clibrate_singal_met(vin_ch, sh.pre_vin)
    # after vin calibration, the v_res_temp will be the last vin value
    # assign to related record
    sh.sh_org_tab.range((10, 10)).value = lo.atof(v_res_temp)

    v_res_temp = met1.mea_v()
    time.sleep(wait_time)
    sh.sh_org_tab.range((10, 6)).value = lo.atof(v_res_temp)
    load1.chg_out(sh.iload2_set, sh.loa_ch2_set, 'off')

else:
    # update the pre-test double check window, see if the pre-test works
    if sh.pre_test_en == 1:
        # pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.pwr_ch_set, 'on')
        print('pre-power on here')

        msg_res = win32api.MessageBox(0, 'press enter if hardware configuration is correct',
                                      'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

    print('pre-power on state finished and ready for next')
    sh.sh_org_tab.range((10, 4)).value = '3.31'
    sh.sh_org_tab.range((10, 6)).value = '3.305'
    # input()

    # == AVDD measurement end

    # == EL-power measurement


x_swire = 0
# counter for SWIRE pulse amount
while x_swire < sh.c_swire:

    # update the MCU pulse first and wait small time

    # ideal V decide the channel of relay board
    ideal_v = sh.ideal_v_table(x_swire)
    if ideal_v > 0:
        meter_ch_ctrl = 2
    if ideal_v < 0:
        meter_ch_ctrl = 3
    print(meter_ch_ctrl)

    if sim_mcu == 1:

        uart_cmd_str = (chr(5) + mcu_cmd_arry[int(meter_ch_ctrl)])
        print(uart_cmd_str)
        mcu_com.write(uart_cmd_str)
        time.sleep(wait_small)
        # input()
        # finished adjust relay channel

        # MCU update SWIRE pulse

        # SWIRE command for the maximum output voltage of ELVDD and ELVSS
        # SWIRE default status need to be high
        pulse1 = sh.sh_org_tab.range((11 + x_swire, 2)).value
        pulse2 = sh.sh_org_tab.range((11 + x_swire, 8)).value

        uart_cmd_str = chr(mode) + chr(int(pulse1)) + chr(int(pulse2))
        print(uart_cmd_str)
        mcu_com.write(uart_cmd_str)
        time.sleep(wait_time)
        # input()

    else:
        # items check in simulation mode
        uart_cmd_str = (chr(5) + mcu_cmd_arry[int(meter_ch_ctrl)])
        print(uart_cmd_str)

        pulse1 = sh.sh_org_tab.range((11 + x_swire, 2)).value
        pulse2 = sh.sh_org_tab.range((11 + x_swire, 8)).value

        uart_cmd_str = chr(mode) + chr(int(pulse1)) + chr(int(pulse2))
        print(uart_cmd_str)
        print('the pulse is ' + str(pulse1) + ' ' + str(pulse2))
        # input()

    if sim_real == 1:

        # call vin calibration
        vin_clibrate_singal_met(vin_ch, sh.pre_vin)
        # after vin calibration, the v_res_temp will be the last vin value
        # assign to related record
        sh.sh_org_tab.range((11 + x_swire, 9)).value = lo.atof(v_res_temp)

        time.sleep(wait_time)
        # measurement start after the SWIRE pulse is set properly
        v_res_temp = met1.mea_v()
        time.sleep(wait_time)
        sh.sh_org_tab.range((11 + x_swire, 4)).value = lo.atof(v_res_temp)

        load1.chg_out(sh.iload1_set, sh.loa_ch_set, 'on')
        time.sleep(wait_time)

        # call vin calibration
        vin_clibrate_singal_met(vin_ch, sh.pre_vin)
        # after vin calibration, the v_res_temp will be the last vin value
        # assign to related record
        sh.sh_org_tab.range((11 + x_swire, 10)).value = lo.atof(v_res_temp)

        time.sleep(wait_time)
        v_res_temp = met1.mea_v()
        time.sleep(wait_time)
        sh.sh_org_tab.range((11 + x_swire, 6)).value = lo.atof(v_res_temp)
        load1.chg_out(sh.iload1_set, sh.loa_ch_set, 'off')

    else:
        # finished the measurement at simulation mode
        print('finished the simulation mode at measurement')
        sh.sh_org_tab.range((11 + x_swire, 4)).value = x_swire
        sh.sh_org_tab.range((11 + x_swire, 6)).value = x_swire + 1
        # simulation mode result

    # save the result after each counter finished
    sh.wb_res.save(sh.result_book_trace)

    x_swire = x_swire + 1

# after the program is finished, back to turn off

mcu_com.write(chr(5) + '00')
print(chr(5) + '00')
time.sleep(wait_time)

# turn off load and power supply
pwr1.change_V(0)
# only turn off the power supply channel but not the relay
load1.chg_out(0, sh.loa_ch_set, 'off')
load1.chg_out(0, sh.loa_ch2_set, 'off')

print('finsihed and goodbye')
