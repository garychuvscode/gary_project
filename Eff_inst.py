# Efficiency test (SWIRE or I2C and others)
# general format of the efficiency testing
# tthis program used to record the efficiency of IC
# testing will need to build from single channel to 3-channel
# from 2-meter only to 5 meter, also consider to add the source meter testing,
# since the loading current control of source meter seems to be better than chromal load
# for the wearable application (low current comparison with chroma CCL mode)

# 20220322 => single channel eff without chart, no source meter, 2-met version
# 20220429 => change the include of inst_pkg to b version, which include the object of source meter
# source meter can be one of choice of load for low current application


# general import package

from ssl import CHANNEL_BINDING_TYPES
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
import sheet_ctrl_eff_inst as sh
# loading of information of contorol book and parameter
import inst_pkg_c as inst
# loading for instrument object ()
# controlled by simulation mode variable

import locale as lo
# include for atof function => transfer string to float

wait_time = sh.wait_time
wait_small = 0.2
wait_sim = 0.02  # waiting time for simulation mode test
# general waiting time for sleep function

# temp usage srting, don't used for global variable
extra_file_name = ''

# temp sheet name for the plot index
# because there are different mode, no need specific channel name, just the positive and negative
eff_temp = ''
pos_temp = ''
neg_temp = ''
raw_temp = ''
# 220825 add for vout and von regulation
pos_pre_temp = ''
neg_pre_temp = ''

y_raw_start = sh.raw_y_position_start
x_raw_start = sh.raw_x_position_start

v_res_temp = 0
# measurement result temp variable
# add the instrument control channel global variable for the result of each channel
v_res_pwr_ch1 = 0
v_res_pwr_ch2 = 0
# v_res_pwr_ch is used for channel 1 and 2, 3 is used for the suto testing


pulse1 = 0
pulse2 = 0
pmic_mode = 4
# mode sequence: 0-3: (EN, SW) = (0, 0),  (0, 1), (1, 0), (1, 1) => default normal

# Vin status global variable
vin_status = ''
# I_AVDD status global variable
i_avdd_status = ''
# I_EL staatus global variable
i_el_status = ''
# SW_I2C status global variable
sw_i2c_status = ''

# i2c register and data (single byte data); slave address is fixed in the MCU
# here only support for the register and data change
reg_i2c = ''
data_i2c = ''

# program status temp string
pro_status_str = ''

bypass_measurement_flag = 0
# when flag = 1, bypass the measurement

# chroma offset current calibration
# by using the result of no load as calibration
value_i_offset1 = 0
value_i_offset2 = 0

sim_real = 0
# 1 => real mode, 0=> simulation(program debug) mode
sim_mcu = 0
# this is only used for MCU com port test simuation
# when setting (sim_real, sim_mcu) = (0, 1)
# enable MCU testing when GPIC disable
# because MCU will be separate with GPIB for implementation and test
mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']
# array mpaaing for the relay control
meter_ch_ctrl = 0
# meter channel indicator: 0: Vin, 1: AVDD, 2: OVDD, 3: OVSS, 4: VOP, 5: VON

# initialization the temp saving parameter for the efficiency calculation
# clear result for each round of the current loop
value_elvdd = 0
value_elvss = 0
value_avdd = 0
value_iin = 0
value_vin = 0
value_iel = 0
value_iavdd = 0
value_eff = 0


eff_done = 0
# the eff done is flag that auto test is already finished one round and wait for reset


# control variable added at the channel_mode in control sheet
# decide now is operate 1(AVDD only), 2(EL_power) or 3(EL_power + AVDD) channel efficiency
# mapping function: 0 => EL_power only; 1 => AVDD only; 2=> 3 channel

# 3 channel mapped to the real condition, define one AVDD voltage
# then change the loading of EL_power for total system efficiency
# table format refer to the define of the excel sheet

# different mode used in the operation
mcu_mode_swire = 1
mcu_mode_sw_en = 3
mcu_mode_I2C = 4
mcu_mode_8_bit_IO = 5
mcu_mode_pat_gen_py = 6
mcu_mode_pat_gen_encode = 7
mcu_mode_pat_gen_direct = 8
# MCU mapping for different mode control in 2553
# MCU mapping for RA_GPIO control is mode 5
# both mode 1 and 2 should defined as dual SWIRE, need to send two pulse command at one time
# need to be build in the control sheet

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
# for efficiency
raw_active = ''
# for raw data
vout_p_active = ''
# for ELVDD, or AVDD
vout_n_active = ''
# for ELVSS
vout_p_pre_active = ''
# for ELVDD, or AVDD
vout_n_pre_active = ''
# for ELVSS


# the excel table gap for the data in raw sheet
raw_gap = 4 + 10
# setting of raw gap is the " gap + element "
# single eff: Vin, Iin, Vout, Iout, Eff => 5 elements

sim_v_data_temp = 0

# ===== sub program definition


def check_inst_update():
    # first load the refesh parameter from the control sheet
    sh.check_refresh()

    pwr_refresh = sh.insctl_pwr_refresh
    load_refresh = sh.insctl_load_refresh
    met1_refresh = sh.insctl_met1_refresh
    met2_refresh = sh.insctl_met2_refresh
    src_refresh = sh.insctl_src_refresh

    # 0522 add this part to keep update the control parameter
    # when every time call the sub program
    sh.program_exit = sh.sh_main.range('B12').value
    # program exit will check every loop finished and break until stop if set ot 0

    sh.inst_auto_selection = sh.sh_main.range('B11').value

    # pwr refresh
    if pwr_refresh == 0:
        # refresh every time enter the program, run refresh process
        sh.para_update_pwr()
        # transfer the on and off command to related string

        # channel 1 update and control
        if sh.insctl_pwr_outs_ch1 == 0:
            # 220514: need to transfer command 0 and 1 to the command need for the instrument
            temp_status = 'off'
        elif sh.insctl_pwr_outs_ch1 == 1:
            temp_status = 'on'
        if sim_real == 1:
            # only give command to the instrument when real mode enable
            pwr1.chg_out(sh.insctl_pwr_vset_ch1,
                         sh.insctl_pwr_iset_ch1, sh.inst_pwr_ch1, temp_status)
            print('PWR change ch1, vset: ' + str(sh.insctl_pwr_vset_ch1) + '; iset: ' +
                  str(sh.insctl_pwr_iset_ch1) + '; status: ' + str(temp_status))
            if sh.insctl_pwr_outs_ch1 == 1:
                # if the power channel is turned on, need the Vin calibration
                if sh.insctl_pwr_calibration == 1:
                    # need to have vin calibration
                    vin_clibrate_singal_met(6, sh.insctl_pwr_vset_ch1)
                sh.pwr_v_ch1_status = v_res_pwr_ch1
                sh.pwr_i_ch1_status = pwr1.read_iout(sh.inst_pwr_ch1)
                pass
            else:
                # set the status of ouptut voltage to 0
                # for the unused channel of power supply
                sh.pwr_v_ch1_status = 0
                sh.pwr_i_ch1_status = 0

            sh.pwr_o_ch1_status = temp_status
            # updatet the on and off status here

            pass
        elif sim_real == 0:
            # run the testing of simulation mode of update
            # just check the variable from the terminal
            print('now is sim mode of pwr update')
            print('PWR change ch1, vset: ' + str(sh.insctl_pwr_vset_ch1) + '; iset: ' +
                  str(sh.insctl_pwr_iset_ch1) + '; status: ' + str(temp_status))
            # also set the status to equal to input settings
            sh.pwr_v_ch1_status = sh.insctl_pwr_vset_ch1
            sh.pwr_i_ch1_status = sh.insctl_pwr_iset_ch1
            sh.pwr_o_ch1_status = temp_status

            pass
        # channel 1 end ======

        # channel 2 update and control
        if sh.insctl_pwr_outs_ch2 == 0:
            # 220514: need to transfer command 0 and 1 to the command need for the instrument
            temp_status = 'off'
        elif sh.insctl_pwr_outs_ch2 == 1:
            temp_status = 'on'
        if sim_real == 1:
            # only give command to the instrument when real mode enable
            pwr1.chg_out(sh.insctl_pwr_vset_ch2,
                         sh.insctl_pwr_iset_ch2, sh.inst_pwr_ch2, temp_status)
            print('PWR change ch2, vset: ' + str(sh.insctl_pwr_vset_ch2) + '; iset: ' +
                  str(sh.insctl_pwr_iset_ch2) + '; status: ' + str(temp_status))
            if sh.insctl_pwr_outs_ch2 == 1:
                # if the power channel is turned on, need the Vin calibration
                if sh.insctl_pwr_calibration == 1:
                    # need to have vin calibration
                    vin_clibrate_singal_met(7, sh.insctl_pwr_vset_ch2)
                sh.pwr_v_ch2_status = v_res_pwr_ch2
                sh.pwr_i_ch2_status = pwr1.read_iout(sh.inst_pwr_ch2)
                pass
            else:
                # set the status of ouptut voltage to 0
                # for the unused channel of power supply
                sh.pwr_v_ch2_status = 0
                sh.pwr_i_ch2_status = 0

            sh.pwr_o_ch2_status = temp_status
            # updatet the on and off status here

            pass
        elif sim_real == 0:
            # run the testing of simulation mode of update
            # just check the variable from the terminal
            print('now is sim mode of pwr update')
            print('PWR change ch2, vset: ' + str(sh.insctl_pwr_vset_ch2) + '; iset: ' +
                  str(sh.insctl_pwr_iset_ch2) + '; status: ' + str(temp_status))
            # also set the status to equal to input settings
            sh.pwr_v_ch2_status = sh.insctl_pwr_vset_ch2
            sh.pwr_i_ch2_status = sh.insctl_pwr_iset_ch2
            sh.pwr_o_ch2_status = temp_status

            pass
        # channel 2 end ======

        # channel 3 update and control
        # ch3 is used by auto test, bypass if auto test is already enable
        if sh.inst_auto_selection == 1:
            sh.pwr_auto_ch3_status = 1
            if sh.insctl_pwr_outs_ch3 == 0:
                # 220514: need to transfer command 0 and 1 to the command need for the instrument
                temp_status = 'off'
            elif sh.insctl_pwr_outs_ch3 == 1:
                temp_status = 'on'
            if sim_real == 1:
                # only give command to the instrument when real mode enable
                pwr1.chg_out(sh.insctl_pwr_vset_ch3,
                             sh.insctl_pwr_iset_ch3, sh.inst_pwr_ch3, temp_status)
                print('PWR change ch3, vset: ' + str(sh.insctl_pwr_vset_ch3) + '; iset: ' +
                      str(sh.insctl_pwr_iset_ch3) + '; status: ' + str(temp_status))
                if sh.insctl_pwr_outs_ch3 == 1:
                    # if the power channel is turned on, need the Vin calibration
                    if sh.insctl_pwr_calibration == 1:
                        # need to have vin calibration
                        vin_clibrate_singal_met(0, sh.insctl_pwr_vset_ch3)
                    sh.pwr_v_ch3_status = v_res_temp
                    sh.pwr_i_ch3_status = pwr1.read_iout(
                        sh.inst_pwr_ch3)
                    pass
                else:
                    # set the status of ouptut voltage to 0
                    # for the unused channel of power supply
                    sh.pwr_v_ch3_status = 0
                    sh.pwr_i_ch3_status = 0

                sh.pwr_o_ch3_status = temp_status
                # updatet the on and off status here

                pass
            elif sim_real == 0:
                # run the testing of simulation mode of update
                # just check the variable from the terminal
                print('now is sim mode of pwr update')
                print('PWR change ch3, vset: ' + str(sh.insctl_pwr_vset_ch3) + '; iset: ' +
                      str(sh.insctl_pwr_iset_ch3) + '; status: ' + str(temp_status))
                # also set the status to equal to input settings
                sh.pwr_v_ch3_status = sh.insctl_pwr_vset_ch3
                sh.pwr_i_ch3_status = sh.insctl_pwr_iset_ch3
                sh.pwr_o_ch3_status = temp_status

                pass
            # channel 3 end ======
            pass
        else:
            sh.pwr_auto_ch3_status = 0
            # channel 3 is used from auto test, set to 0 and turn red

        # bypass the serial function first, since there are no serial needed yet
        # only need to change the serial function when line transient control is added to the
        # auto testing, otherwise, no need for the setial finction
        # send the serial control command here for the PWR, but may need to add the control stringin the
        # instpkg file
        # 220516

        # after the setting of power finished, vall the function update status
        sh.status_update_pwr()

        pass
    elif pwr_refresh == 1:
        # at the latch mode, pass directly
        pass
    elif pwr_refresh == 2:
        # follow the general refresh settings
        # decide to skip this function first, build in the future if needed
        # 220514: to complicated, skip too many different refresh fuction, and build other part first
        # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
        # other will check after finished, if real a must have function or not
        pass

    pass

    # loader refresh time
    if load_refresh == 0:
        # refresh every time enter the program, run refresh process
        sh.para_update_load()
        # transfer the on and off command to related string
        # not

        # ch1 update and control
        # byapss the ch1 and ch2 loader if auto testing is enable
        if sh.inst_auto_selection == 1:
            if sh.insctl_load_outs_ch1 == 0:
                # 220514: need to transfer command 0 and 1 to the command need for the instrument
                temp_status = 'off'
            elif sh.insctl_load_outs_ch1 == 1:
                temp_status = 'on'
            if sim_real == 1:
                # only give command to the instrument when real mode enable
                # change the mode first to fit the change of current
                load1.chg_mode(1, sh.insctl_load_mset_ch1)
                # change the output current setting after the mode set change
                load1.chg_out(sh.insctl_load_iset_ch1, 1, temp_status)
                print('load change ch1, iset: ' + str(sh.insctl_load_iset_ch1) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch1))

                sh.load_i_ch1_status = load1.read_iout(1)
                # read the load current and update to the status
                sh.load_m_ch1_status = load1.mode_o[0]
                # check the final mode setting and update to the status

                pass
            elif sim_real == 0:
                # run the testing of simulation mode of update
                # just check the variable from the terminal
                print('now is sim mode of lpwroad update')
                print('load change ch1, iset: ' + str(sh.insctl_load_iset_ch1) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch1))
                sh.load_i_ch1_status = sh.insctl_load_iset_ch1
                sh.load_m_ch1_status = sh.insctl_load_outs_ch1
                pass

            sh.load_o_ch1_status = temp_status
            # update the final output status to the status variable
            pass
        else:
            sh.load_auto_ch1_status = 0
            # channel 1 is used from auto test, set to 0 and turn red

        # ch1 end ======

        # ch2 update and control
        # byapss the ch1 and ch2 loader if auto testing is enable
        if sh.inst_auto_selection == 1:
            if sh.insctl_load_outs_ch2 == 0:
                # 220514: need to transfer command 0 and 1 to the command need for the instrument
                temp_status = 'off'
            elif sh.insctl_load_outs_ch2 == 1:
                temp_status = 'on'
            if sim_real == 1:
                # only give command to the instrument when real mode enable
                # change the mode first to fit the change of current
                load1.chg_mode(2, sh.insctl_load_mset_ch2)
                # change the output current setting after the mode set change
                load1.chg_out(sh.insctl_load_iset_ch2, 2, temp_status)
                print('load change ch2, iset: ' + str(sh.insctl_load_iset_ch2) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch2))

                sh.load_i_ch2_status = load1.read_iout(2)
                # read the load current and update to the status
                sh.load_m_ch2_status = load1.mode_o[1]
                # check the final mode setting and update to the status

                pass
            elif sim_real == 0:
                # run the testing of simulation mode of update
                # just check the variable from the terminal
                print('now is sim mode of load update')
                print('load change ch2, iset: ' + str(sh.insctl_load_iset_ch2) + '; status: ' +
                      str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch2))

                sh.load_i_ch2_status = sh.insctl_load_iset_ch2
                sh.load_m_ch2_status = sh.insctl_load_outs_ch2

                pass

            sh.load_o_ch2_status = temp_status
            # update the final output status to the status variable

            pass
        else:
            sh.load_auto_ch2_status = 0
            # channel 1 is used from auto test, set to 0 and turn red

        # ch2 end ======

        # ch3 update and control
        if sh.insctl_load_outs_ch3 == 0:
            # 220514: need to transfer command 0 and 1 to the command need for the instrument
            temp_status = 'off'
        elif sh.insctl_load_outs_ch3 == 1:
            temp_status = 'on'
        if sim_real == 1:
            # only give command to the instrument when real mode enable
            # change the mode first to fit the change of current
            load1.chg_mode(3, sh.insctl_load_mset_ch3)
            # change the output current setting after the mode set change
            load1.chg_out(sh.insctl_load_iset_ch3, 3, temp_status)
            print('load change ch3, iset: ' + str(sh.insctl_load_iset_ch3) + '; status: ' +
                  str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch3))

            sh.load_i_ch3_status = load1.read_iout(3)
            # read the load current and update to the status
            sh.load_m_ch3_status = load1.mode_o[2]
            # check the final mode setting and update to the status

            pass
        elif sim_real == 0:
            # run the testing of simulation mode of update
            # just check the variable from the terminal
            print('now is sim mode of load update')
            print('load change ch3, iset: ' + str(sh.insctl_load_iset_ch3) + '; status: ' +
                  str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch3))

            sh.load_i_ch3_status = sh.insctl_load_iset_ch3
            sh.load_m_ch3_status = sh.insctl_load_outs_ch3

            pass

        sh.load_o_ch3_status = temp_status
        # update the final output status to the status variable

        # ch3 end ======

        # ch4 update and control
        if sh.insctl_load_outs_ch4 == 0:
            # 220514: need to transfer command 0 and 1 to the command need for the instrument
            temp_status = 'off'
        elif sh.insctl_load_outs_ch4 == 1:
            temp_status = 'on'
        if sim_real == 1:
            # only give command to the instrument when real mode enable
            # change the mode first to fit the change of current
            load1.chg_mode(4, sh.insctl_load_mset_ch4)
            # change the output current setting after the mode set change
            load1.chg_out(sh.insctl_load_iset_ch4, 4, temp_status)
            print('load change ch4, iset: ' + str(sh.insctl_load_iset_ch4) + '; status: ' +
                  str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch4))

            sh.load_i_ch4_status = load1.read_iout(4)
            # read the load current and update to the status
            sh.load_m_ch4_status = load1.mode_o[3]
            # check the final mode setting and update to the status

            pass
        elif sim_real == 0:
            # run the testing of simulation mode of update
            # just check the variable from the terminal
            print('now is sim mode of load update')
            print('load change ch4, iset: ' + str(sh.insctl_load_iset_ch4) + '; status: ' +
                  str(temp_status) + '; mode set:' + str(sh.insctl_load_mset_ch4))

            sh.load_i_ch4_status = sh.insctl_load_iset_ch4
            sh.load_m_ch4_status = sh.insctl_load_outs_ch4

            pass

        sh.load_o_ch4_status = temp_status
        # update the final output status to the status variable

        # ch4 end ======

        pass
    elif load_refresh == 1:
        # at the latch mode, pass directly
        pass
    elif load_refresh == 2:
        # follow the general refresh settings
        # decide to skip this function first, build in the future if needed
        # 220514: to complicated, skip too many different refresh fuction, and build other part first
        # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
        # other will check after finished, if real a must have function or not
        pass

    sh.status_update_load()
    # change the status after the parameter update finished

    # meter1 refresh time
    # add inst_auto selection => when only inst active meter control
    if met1_refresh == 0:
        sh.para_update_met1()
        # update the other parameter after refresh checked
        # then start the other parameter update later
        if sh.inst_auto_selection == 1:
            # update meter connection to 1 if available
            sh.met1_connection_status = 1

            if sh.insctl_met1_mset == 0:
                # enter the voltage measurement mode
                if sim_real == 1:
                    sh.met1_v_mea_status = met1.mea_v2(0, sh.insctl_met1_leve)
                    time.sleep(wait_small)
                    sh.met1_i_mea_status = 'NA'
                    sh.met1_mode_status = 'votlage'
                    sh.met1_level_status = sh.insctl_met1_leve
                    pass
                else:
                    print('meter 1 measure voltage')
                    sh.met1_v_mea_status = 'sim_volt1'
                    pass
                pass

            elif sh.insctl_met1_mset == 1:
                # enter the current measurement mode
                if sim_real == 1:

                    sh.met1_i_mea_status = met1.mea_i2(0, sh.insctl_met1_leve)
                    time.sleep(wait_small)
                    sh.met1_v_mea_status = 'NA'
                    sh.met1_mode_status = 'current'
                    sh.met1_level_status = sh.insctl_met1_leve
                    pass
                else:
                    print('meter 1 measure current')
                    sh.met1_i_mea_status = 'sim_curr1'
                    pass
                pass

            pass
        else:
            # meter channel is decide by auto control
            sh.met1_connection_status = 0
            # 0523 other keep in the initial value
    elif met1_refresh == 1:
        # at the latch mode, pass directly
        pass
    elif met1_refresh == 2:
        # follow the general refresh settings
        # decide to skip this function first, build in the future if needed
        # 220514: to complicated, skip too many different refresh fuction, and build other part first
        # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
        # other will check after finished, if real a must have function or not
        pass

    sh.status_update_met1()

    # meter2 refresh time
    # add inst_auto selection => when only inst active meter control
    if met2_refresh == 0:
        sh.para_update_met2()
        # update the other parameter after refresh checked
        # then start the other parameter update later
        if sh.inst_auto_selection == 1:
            sh.met2_connection_status = 1

            if sh.insctl_met2_mset == 0:
                # enter the voltage measurement mode
                if sim_real == 1:
                    sh.met2_v_mea_status = met2.mea_v2(0, sh.insctl_met2_leve)
                    time.sleep(wait_small)
                    sh.met2_i_mea_status = 'NA'
                    sh.met2_mode_status = 'votlage'
                    sh.met2_level_status = sh.insctl_met2_leve
                    pass
                else:
                    print('meter 1 measure voltage')
                    sh.met2_v_mea_status = 'sim_volt2'
                    pass
                pass

            elif sh.insctl_met2_mset == 1:
                # enter the current measurement mode
                if sim_real == 1:

                    sh.met2_i_mea_status = met2.mea_i2(0, sh.insctl_met2_leve)
                    time.sleep(wait_small)
                    sh.met2_v_mea_status = 'NA'
                    sh.met2_mode_status = 'current'
                    sh.met2_level_status = sh.insctl_met2_leve
                    pass
                else:
                    print('meter 1 measure current')
                    sh.met2_i_mea_status = 'sim_curr2'
                    pass
                pass
            pass
        else:
            # meter channel is decide by auto control
            # 0523 other keep in the initial value
            sh.met2_connection_status = 0

    elif met2_refresh == 1:
        # at the latch mode, pass directly
        pass
    elif met2_refresh == 2:
        # follow the general refresh settings
        # decide to skip this function first, build in the future if needed
        # 220514: to complicated, skip too many different refresh fuction, and build other part first
        # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
        # other will check after finished, if real a must have function or not
        pass

    sh.status_update_met2()

    #  source meter refresh time
    # add inst_auto selection => when only inst active source meter control
    if src_refresh == 0:
        sh.para_update_src()
        # update the other parameter after refresh checked
        # then start the other parameter update later
        if sh.inst_auto_selection == 1:
            sh.src_connection_status = 1
            # change the status to string
            if sh.insctl_src_outs == 1:
                sh.insctl_src_outs = 'on'
            elif sh.insctl_src_outs == 0:
                sh.insctl_src_outs = 'off'

            if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:

                if sh.insctl_src_mset == 0:
                    # enter the current source mode
                    sh.insctl_src_mset = 'CURR'
                    if sim_real == 1:

                        if load_src.source_type_o != sh.insctl_src_mset:
                            # the only difference is need to turn off load before change mode
                            load_src.load_off()

                        # if the last output state of source meter is same in current source
                        # no need to change type, just change output
                        load_src.change_type(
                            sh.insctl_src_mset, sh.insctl_src_cset)
                        # if the type doesn't change in the source meter, only change the clamp settings
                        load_src.change_I(sh.insctl_src_leve,
                                          sh.insctl_src_outs)
                        # after sending the command to instrument, read the instruemnt status can know
                        # if the command is allowable, command clamp setting should be build in
                        # inst_pkg side, not the main program side

                        # read status and update to the status output blank in inst control sheet
                        sh.src_mode_status = load_src.source_type_o
                        sh.src_clamp_status = load_src.clamp_VI_o
                        sh.src_level_status = load_src.iset_o
                        sh.src_v_mea_status = load_src.read('VOLT')
                        sh.src_i_mea_status = load_src.read('CURR')
                        sh.src_o_status = load_src.state_o

                        pass
                    else:
                        # read status and update to the status output blank in inst control sheet
                        sh.src_mode_status = sh.insctl_src_mset
                        sh.src_clamp_status = sh.insctl_src_cset
                        sh.src_level_status = sh.insctl_src_leve
                        sh.src_v_mea_status = "load_src.read('VOLT')"
                        sh.src_i_mea_status = "load_src.read('CURR')"
                        sh.src_o_status = sh.insctl_src_outs
                        pass
                    pass

                elif sh.insctl_src_mset == 1:
                    # enter the voltage source mode
                    sh.insctl_src_mset = 'VOLT'
                    if sim_real == 1:
                        if load_src.state_o != sh.insctl_src_mset:
                            # the only difference is need to turn off load before change mode
                            load_src.load_off()

                        # if the last output state of source meter is same in current source
                        # no need to change type, just change output
                        load_src.change_type(
                            sh.insctl_src_mset, sh.insctl_src_cset)
                        # if the type doesn't change in the source meter, only change the clamp settings
                        load_src.change_V(sh.insctl_src_leve,
                                          sh.insctl_src_outs)
                        # after sending the command to instrument, read the instruemnt status can know
                        # if the command is allowable, command clamp setting should be build in
                        # inst_pkg side, not the main program side

                        # read status and update to the status output blank in inst control sheet
                        sh.src_mode_status = load_src.source_type_o
                        sh.src_clamp_status = load_src.clamp_VI_o
                        sh.src_level_status = load_src.iset_o
                        sh.src_v_mea_status = load_src.read('VOLT')
                        sh.src_i_mea_status = load_src.read('CURR')
                        sh.src_o_status = load_src.state_o

                    else:
                        # read status and update to the status output blank in inst control sheet
                        sh.src_mode_status = sh.insctl_src_mset
                        sh.src_clamp_status = sh.insctl_src_cset
                        sh.src_level_status = sh.insctl_src_leve
                        sh.src_v_mea_status = "load_src.read('VOLT')"
                        sh.src_i_mea_status = "load_src.read('CURR')"
                        sh.src_o_status = sh.insctl_src_outs
                        pass
                    pass
                else:
                    # since it's not set to voltage or current source
                    # no action for this case
                    pass
                pass
            pass
        else:
            # set SRC status to 0 if is't control by auto testing
            sh.src_connection_status = 0
            # 0523 other keep in the initial value
            pass
    elif src_refresh == 1:
        # at the latch mode, pass directly
        pass
    elif src_refresh == 2:
        # follow the general refresh settings
        # decide to skip this function first, build in the future if needed
        # 220514: to complicated, skip too many different refresh fuction, and build other part first
        # refresh function only polling by function call( refresh = 0 ) or latch ( refresh = 1 )
        # other will check after finished, if real a must have function or not
        pass

    sh.status_update_src()

    pass


def mcu_write(index):
    # the MCU write used to generate the command string and send the command out
    # but the related control parameter need to define in the main program
    # here is only used to reduce the code of generate the string
    # and the MCU UART sending command

    if index == 'swire':
        uart_cmd_str = (chr(1) + chr(int(pulse1)) + chr(int(pulse2)))
        # for the SWIRE mode of 2553, there are 2 pulse send to the MCU and DUT
        # pulse amount is from 1 to 255, not sure if 0 will have error or not yet
        # 20220121
    elif index == 'en_sw':
        uart_cmd_str = (chr(3) + chr(int(pmic_mode)) + chr(1))
        # for the EN SWIRE control mode, need to handle the recover to normal mode (EN, SW) = (1, 1)
        # at the end of application
        # this mode only care about the first data ( 0-4 )
    elif index == 'relay':
        uart_cmd_str = (chr(5) + mcu_cmd_arry[meter_ch_ctrl])
        # assign relay to related channel after function called
        # channel index is from golbal variable

    elif index == 'i2c':
        uart_cmd_str = (chr(4) + str(reg_i2c) + str(data_i2c))
        # send mapped i2c command out from MCU

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

    # define the globa variable for eff calculation
    global value_elvdd
    global value_elvss
    global value_avdd
    global value_iin
    global value_vin
    global value_iel
    global value_iavdd

    global bypass_measurement_flag
    # first to check if the bypass flag raise ~
    # set measurement result to 0 if the bpass flag is enable
    if bypass_measurement_flag == 1:
        mea_res = '0'

    if data_name == 'vin':
        # vin only record in the raw data
        raw_active.range((11 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        value_vin = float(mea_res)

    elif data_name == 'iin':
        # iin only record in the raw data
        raw_active.range((12 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        value_iin = float(mea_res)

    elif data_name == 'elvdd':
        # elvdd record in the raw data, elvdd regulation
        raw_active.range((13 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        # sheet_active.range((25 + x_iload, 3 + x_vin)).value = lo.atof(mea_res)
        if sh.channel_mode == 0 or sh.channel_mode == 2:
            vout_p_active.range((25 + x_iload, 3 + x_vin)
                                ).value = lo.atof(mea_res)
        value_elvdd = float(mea_res)

    elif data_name == 'elvss':
        # elvss record in the raw data, elvss regulation
        raw_active.range((14 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        if sh.channel_mode == 0 or sh.channel_mode == 2:
            vout_n_active.range((25 + x_iload, 3 + x_vin)
                                ).value = lo.atof(mea_res)
        value_elvss = float(mea_res)

    elif data_name == 'i_el':
        # i_el only record in the raw data
        raw_active.range((15 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res) - value_i_offset1
        value_iel = float(mea_res) - value_i_offset1

    elif data_name == 'avdd':
        # avvdd record in the raw data, avvdd regulation
        raw_active.range((16 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        value_avdd = float(mea_res)
        # the raw data of AVDD need to record no matter in AVDD only mode or the 3-ch mode
        # the selection of EL only or not is decidde inthe main program
        # here is only for the choice of AVDD regulation
        if sh.channel_mode == 1:
            # only need to record the regulation when is operating for AVDD only mode
            vout_p_active.range((25 + x_iload, 3 + x_vin)
                                ).value = lo.atof(mea_res)

    elif data_name == 'i_avdd':
        # i_avdd only record in the raw data
        raw_active.range((17 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res) - value_i_offset2
        value_iavdd = float(mea_res) - value_i_offset2

    elif data_name == 'eff':
        # eff record in the raw data
        raw_active.range((18 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        sheet_active.range((25 + x_iload, 3 + x_vin)
                           ).value = lo.atof(mea_res)

    elif data_name == 'vop':
        # vop record in the raw data
        raw_active.range((19 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        vout_p_pre_active.range((25 + x_iload, 3 + x_vin)
                                ).value = lo.atof(mea_res)

    elif data_name == 'von':
        # von record in the raw data
        raw_active.range((20 + raw_gap * x_vin, 3 + x_iload)
                         ).value = lo.atof(mea_res)
        vout_n_pre_active.range((25 + x_iload, 3 + x_vin)
                                ).value = lo.atof(mea_res)

    # clear the bypass flag every time enter data latch function
    bypass_measurement_flag = 0


# definition of the program statue update to the excel main sheet
def program_status(status_string):
    # transfer to the string for following operation
    # if you need to modify in sub-program, need to use global definition
    # global vin_status
    # global i_avdd_status
    # global i_el_status
    # global sw_i2c_status
    status_sting_sub = str(status_string)
    sh.sh_main.range((3, 2)).value = status_sting_sub
    # Vin status
    sh.sh_main.range('F3').value = vin_status
    # I_AVDD status
    sh.sh_main.range('F4').value = i_avdd_status
    # I_EL staatus
    sh.sh_main.range('F5').value = i_el_status
    # SW_I2C status
    sh.sh_main.range('F6').value = sw_i2c_status

    print('status_update: ' + status_sting_sub)
    print(str(vin_status) + '-' + str(i_avdd_status) +
          '-' + str(i_el_status) + '-' + str(sw_i2c_status))
    # use for debugging for the program status update
    # input()

# definition of the calibration sub program


def vin_clibrate_singal_met(vin_ch, vin_target):
    # vin_ch is the setting for relay, which means: 0, 6, 7
    # it will mapped to the provide channel setting in the excel

    global v_res_temp
    global v_res_pwr_ch1
    global v_res_pwr_ch2

    # definition of Vin channel => for the auto tesing default Vin channel is set to 0
    # other Vin channel for Vin control is set to 6 and 7
    # after consider the auto Vin(0), 5 output channel(1-5), other 2 Vin channel will be
    # relay channel 6 and 7

    vin_diff = 0
    v_res_temp_f = 0
    # 20220103 add the Vin adjustment function
    # change the measure channel to the Vin channel (No.3)
    # this part added after the load turned on for better Vin setting
    mcu_com.write(chr(5) + mcu_cmd_arry[vin_ch])
    print(chr(5) + mcu_cmd_arry[vin_ch])
    time.sleep(wait_time)
    # measure the first Vin after relay change
    if vin_ch == 0:
        v_res_temp = met1.mea_v()
        v_res_temp_f = lo.atof(v_res_temp)
        # the Vin calibration starts from here
        pass

    elif vin_ch == 6:
        v_res_pwr_ch1 = met1.mea_v()
        v_res_temp_f = lo.atof(v_res_pwr_ch1)
        # the Vin calibration starts from here
        pass

    elif vin_ch == 7:
        v_res_pwr_ch2 = met1.mea_v()
        v_res_temp_f = lo.atof(v_res_pwr_ch2)
        # the Vin calibration starts from here
        pass

    vin_diff = vin_target - v_res_temp_f
    vin_new = vin_target
    while vin_diff > sh.vin_diff_set or vin_diff < (-1 * sh.vin_diff_set):
        vin_new = vin_new + 0.5 * (vin_target - v_res_temp_f)
        # clamp for the Vin maximum
        if vin_new > sh.pre_vin_max:
            vin_new = sh.pre_vin_max

        if vin_ch == 0:
            pwr1.change_V(vin_new, sh.inst_pwr_ch3)
            # send the new Vin command for the auto testing channel
            pass

        elif vin_ch == 6:
            pwr1.change_V(vin_new, sh.inst_pwr_ch1)
            # change the vsetting of channel 1 (mapped in program)
            pass

        elif vin_ch == 7:
            pwr1.change_V(vin_new, sh.inst_pwr_ch2)
            # change the vsetting of channel 2 (mapped in program)
            pass

        time.sleep(wait_time)
        # # measure the Vin change result
        # v_res_temp = met1.mea_v()
        # v_res_temp_f = lo.atof(v_res_temp)

        # 220524: update the calibration result to status variable

        if vin_ch == 0:
            v_res_temp = met1.mea_v()
            v_res_temp_f = lo.atof(v_res_temp)
            # the Vin calibration starts from here
            pass

        elif vin_ch == 6:
            v_res_pwr_ch1 = met1.mea_v()
            v_res_temp_f = lo.atof(v_res_pwr_ch1)
            # the Vin calibration starts from here
            pass

        elif vin_ch == 7:
            v_res_pwr_ch2 = met1.mea_v()
            v_res_temp_f = lo.atof(v_res_pwr_ch2)
            # the Vin calibration starts from here
            pass

        vin_diff = vin_target - v_res_temp_f

    # after the loop is finished, record the Vin meausred result (for full load) at the end
    # sh.sh_org_tab.range((10, 9)).value = lo.atof(v_res_temp)
    # the Vin calibration ends from here

    # after the Vin calibration is finished, change the measure channel back
    mcu_com.write(chr(5) + mcu_cmd_arry[meter_ch_ctrl])
    print(chr(5) + mcu_cmd_arry[meter_ch_ctrl])
    time.sleep(wait_time)
    # finished getting back to the initial state


# ===== sub program finished


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

    # MCU initial for EN, SW under normal mode (EN, SW control is at mode 3)
    uart_cmd_str = chr(3) + chr(4) + chr(1)
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)
    time.sleep(wait_time)
    # input()

# MCU need to cover the change from positive to negative
# ELVDD, ELVSS and AVDD channel change
# MCU channel, initial setting:

if sim_real == 1:
    # start the general instrument startup
    # import inst_pkg as inst
    pwr1 = inst.LPS_505N(0, 0, 1, sh.pwr_sup_addr, 'off')
    # no worries about the channel of supply, will change at initial state below

    met1 = inst.Met_34460(sh.met_vin_res, sh.met_vin_rang,
                          sh.met_iin_res, sh.met_iin_rang, sh.met_vdd_addr)
    # met1
    # meter loaded from the main sheet
    met2 = inst.Met_34460(sh.met_vin_res, sh.met_vin_rang,
                          sh.met_iin_res, sh.met_iin_rang, sh.met_vss_addr)

    # 20220429, define the load anyway, but choose to use or not during operation
    load1 = inst.chroma_63600(1, sh.loa_dts_addr, str(sh.loa_mod_set))
    # load channel and related setting also change after definition

    if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
        load_src = inst.Keth_2440(
            0, 0, sh.loa_src_addr, 'off', sh.loader_source_type, sh.v_clamp_load)
    else:
        # there is no source meter needed for the output
        pass
    # loader channel mapping: ch1 is for EL_power and ch2 is for AVDD

    # turn on the GPIB connection of instrument
    pwr1.open_inst()
    met1.open_instr()
    load1.open_inst()
    met2.open_instr()
    if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
        # if not going to open the GPIB device, should be aboe to operate the system
        # without source meter
        load_src.open_inst()
    else:
        # there is no source meter needed for the output
        pass

    # # load the name to related blank in result book
    # sh.sh_main.range('D27').value = pwr1.inst_name()
    # sh.sh_main.range('D28').value = met1.inst_name()
    # sh.sh_main.range('D30').value = load1.inst_name()

    # the other way for update instrument name(sub program in sh_ctrl)
    sh.inst_name_sheet('PWR1', pwr1.inst_name())
    sh.inst_name_sheet('MET1', met1.inst_name())
    sh.inst_name_sheet('MET2', met2.inst_name())
    sh.inst_name_sheet('LOAD1', load1.inst_name())
    if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
        sh.inst_name_sheet('LOADSR', load_src.inst_name())
    else:
        # there is no source meter needed for the output
        pass
    pass
elif sim_real == 0:
    print('initialization of instrument finished')
    pass

# 220521 new start of program, infinite loop for the selection
# choose the eff or the normal instrument control
while (sh.program_exit > 0):
    sh.program_exit = sh.sh_main.range('B12').value
    sh.inst_auto_selection = sh.sh_main.range('B11').value

    # check if the eff_done need to reset before the loop
    # eff_done can only be reset when it's already be 1 to prevent error
    if eff_done == 1:
        sh.eff_rerun()
        eff_done = sh.eff_done_sh

    if sh.inst_auto_selection == 1:
        # 220824 change the condition to prevent stuck in check instrument when running efficiency testing

        # the setting of only instrument control
        check_inst_update()
        # call the update function and refresh the settings
        time.sleep(wait_time)
        # give some delay between each refresh (call subprogram)

        pass
    elif (sh.inst_auto_selection == 0 or sh.inst_auto_selection == 2) and eff_done == 0:
        # 220824 change the condition to prevent stuck in check instrument when running efficiency testing

        # here is the original eff testing program
        # in the deepest loop, will check if auto selection is 0 or 2
        # for 2 it will also call the check instrument and control
        # for 0 is just like the previour efficiency measurement

        # real mode using instrument connection
        if sim_real == 1:
            # 220521: move the insturment initialization to the front
            # settings need for related sheet of instrument

            # protection and other setting (before channel on)

            # power supply OV and OC protection
            pwr1.ov_oc_set(sh.pre_vin_max, sh.pre_imax)

            # power supply channel (channel on setting)
            if sh.pre_test_en == 1:
                pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout,
                             sh.inst_pwr_ch3, 'on')
                print('pre-power on here')
                # turn off the power and load

                load1.chg_out(0, sh.loa_ch_set, 'off')
                load1.chg_out(0, sh.loa_ch2_set, 'off')

                if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
                    load_src.load_off()

                print('also turn all load off')

                if sh.en_start_up_check == 1:
                    msg_res = win32api.MessageBox(
                        0, 'press enter if hardware configuration is correct', 'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

            # the power will change from initial state directly, not turn off between transition

            # should not need the extra Vin in the change
            # pwr1.chg_out(sh.vin1_set, sh.pre_sup_iout, sh.inst_pwr_ch3, 'on')

            # loader channel and current
            # default off, will be turn on and off based on the loop control

            # load1.chg_out(sh.iload1_set, sh.loa_ch_set, 'off')
            # # load set for EL-power
            # load1.chg_out(sh.iload2_set, sh.loa_ch2_set, 'off')
            # # load set for AVDD

            time.sleep(wait_time)

        else:
            # update the pre-test double check window, see if the pre-test works
            if sh.pre_test_en == 1:
                # pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.inst_pwr_ch3, 'on')
                print('pre-power on here')
                if sh.en_start_up_check == 1:
                    msg_res = win32api.MessageBox(0, 'press enter if hardware configuration is correct',
                                                  'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

            print('pre-power on state finished and ready for next')
            # input()
            # can used any input for pre-power on finished test ing on excel
            program_status('test mode power on ok')

            # input()

            # pre-power on test finished here

        # efficiency testing program starts from heere

        # 1st loop is the selection of I2C and SWIRE pulse control

        # selection for the loop control variable
        x_sw_i2c = 0
        c_sw_i2c = 0

        # selection for the SWIRE(1) or I2C(2)
        if sh.sw_i2c_select == 1:
            if sh.channel_mode == 1:
                # if the setting is avdd only ,need to map the counter to AVDD pulse number
                c_sw_i2c = sh.c_avdd_pulse
            elif sh.channel_mode == 0 or sh.channel_mode == 2:
                c_sw_i2c = sh.c_pulse
        elif sh.sw_i2c_select == 2:
            c_sw_i2c = sh.c_i2c
            # i2c group counter setting
            c_i2c_group = sh.c_i2c_g

        # 220823 consider to add the AVDD measurement for both I2C and normal(SWIRE mode)
        # but maybe no need for the

        # error handling can be consider in future, but now should be ok to prevent
        # all the error from knowing the system operating rule and bug

        # # error handling for counter = 0 after data refresh
        # if c_sw_i2c == 0 :
        #     c_sw_i2c = 1
        #     c_single = 1

        while x_sw_i2c < c_sw_i2c:
            # need to set up specific SWIRE pulse setting (2-pulse version) or give the I2C command at this stage
            # before the stage of next loop
            # loaded the SWIRE or I2C command from the control sheet
            # send it out from MCU to testing EVM board

            if sh.sw_i2c_select == 1:
                # SWIRE control loop
                # setup the related 2 pulse
                if sh.channel_mode == 1:
                    # if the setting is avdd only ,need to map the pulse to AVDD pulse
                    pulse1 = sh.sh_org_tab.range((3 + x_sw_i2c, 8)).value
                    pulse2 = pulse1
                    extra_file_name = 'SWIRE_AVDD_'
                elif sh.channel_mode == 0 or sh.channel_mode == 2:
                    # for the SWIRE pulse, need 2 pulse for ELVDD and ELVSS
                    pulse1 = sh.sh_org_tab.range((3 + x_sw_i2c, 5)).value
                    pulse2 = sh.sh_org_tab.range((3 + x_sw_i2c, 6)).value
                    extra_file_name = 'SWIRE_EL_'
                print('pulse1: ' + str(pulse1) +
                      '; and pulse2: ' + str(pulse2) + 'under mode ' + str(sh.channel_mode))
                # send the pulse out through MCU
                if sim_mcu == 1:
                    mcu_write('swire')
                else:
                    print('MCU output function called')

                # call the build file to build new file to save result
                # for the SWIRE pulse control
                extra_file_name = extra_file_name + \
                    str(int(pulse1)) + '_' + str(int(pulse2))
                sw_i2c_status = str(int(pulse1)) + '_' + str(int(pulse2))
                # sh.build_file(str(extra_file_name))
                # pro_status_str = 'file built'
                # program_status(pro_status_str)

            elif sh.sw_i2c_select == 2:
                # setup i2C group counter
                x_i2c_group = 0

                # modify the i2c str before the loop to get all change (register:data)
                # initial the extra_file_name before the loop start
                extra_file_name = 'i2c'
                while x_i2c_group < c_i2c_group:
                    # I2C control loop
                    # set up the i2c related data
                    reg_i2c = sh.sh_org_tab3.range(
                        (3 + c_i2c_group * x_sw_i2c + x_i2c_group, 2)).value
                    data_i2c = sh.sh_org_tab3.range(
                        (3 + c_i2c_group * x_sw_i2c + x_i2c_group, 3)).value
                    print('register: ' + reg_i2c)
                    print('data: ' + data_i2c)
                    if sim_mcu == 1:
                        mcu_write('i2c')
                    else:
                        print('MCU output function called')

                    # after write to the MCU, update the extra name for the file
                    extra_file_name = extra_file_name + '_' + \
                        str(reg_i2c) + '-' + str(data_i2c)
                    sw_i2c_status = str(reg_i2c) + '-' + str(data_i2c)
                    x_i2c_group = x_i2c_group + 1

                # extra_file_name = 'SWIRE_' + str(pulse1) + '_' + str(pulse2)
            sh.build_file(str(extra_file_name))
            pro_status_str = 'file built'
            program_status(pro_status_str)
            print(sh.sh_main)
            print('')
            # 220823 change the file build code to outer loop and save the few lines before

            # finished the swire or i2c command and ready to enter next loop

            # # call the build file to build new file to save result
            # sh.build_file(str(x_sw_i2c))

            x_avdd = 0
            # counter avdd current setting
            # when in channel_mode 0 or 1, there is not option of AVDD current
            # need to add the error handling of AVDD current counter in the loop
            # need to set the aVDD current counter to 1 to pevent of overflow in the sheet array
            c_avdd = 0
            if sh.channel_mode == 0 or sh.channel_mode == 1:
                # for the 1 or 2 channel operation, only used c_avdd once
                # there are no extra loop at this stage
                c_avdd = 1
            else:
                # this is 3-channel operatoion
                c_avdd = sh.c_avdd_load

            while x_avdd < c_avdd:

                # define AVDD current
                curr_avdd = sh.sh_org_tab.range((3 + x_avdd, 4)).value
                print('AVDD current is set to: ' + str(curr_avdd) + ' A')
                if sh.channel_mode == 2 or sh.channel_mode == 0:
                    i_avdd_status = str(curr_avdd)
                    # 20220509 => need to add the (EN, SW) mode setting to the MCU status
                    # since the i2c don't have the mode selection function, don't care
                    # about the PMIC status, need to be config in the register command in
                    # i2c mode
                    if sim_mcu == 1:
                        pmic_mode = 4
                        # set the PMIC to normal mode
                        # and update the MCU write commanad
                        mcu_write('en_sw')
                    else:
                        print('MCU mode is set to (EN, SW) = (1, 1)')
                else:
                    i_avdd_status = 0
                    # since the i2c don't have the mode selection function, don't care
                    # about the PMIC status, need to be config in the register command in
                    # i2c mode
                    if sim_mcu == 1:
                        # 0511: to have better discharge result for PMIC, add the shut down mode between
                        # and then start from AOD mode only
                        pmic_mode = 1
                        mcu_write('en_sw')
                        print('PMIC shut down for EL power discharge')
                        # turn off the PMIC and wait for other channel to discharge
                        time.sleep(wait_small)

                        pmic_mode = 3
                        # set the PMIC to AOD mode
                        # and update the MCU write commanad
                        mcu_write('en_sw')
                    else:
                        print('MCU mode is set to (EN, SW) = (1, 0)')
                pro_status_str = 'AVDD current : ' + str(curr_avdd)
                program_status(pro_status_str)
                # please note that AVDD current need to set to 0 if not using 3-ch mode

                # after generate the related file, should be able to have array for active raw and sheet
                # and the mapping sheet name is change with AVDD current, so it's in AVDD_current loop

                sheet_active = sh.wb_res.sheets(
                    sh.sheet_arry[sh.sub_sh_count * x_avdd])
                eff_temp = str(sh.sheet_arry[sh.sub_sh_count * x_avdd])
                # add the string save sheet name for the usage of plot

                raw_active = sh.wb_res.sheets(
                    sh.sheet_arry[sh.sub_sh_count * x_avdd + 1])
                raw_temp = str(
                    str(sh.sheet_arry[sh.sub_sh_count * x_avdd + 1]))
                # add the string save sheet name for the usage of plot

                vout_p_active = sh.wb_res.sheets(
                    sh.sheet_arry[sh.sub_sh_count * x_avdd + 2])
                pos_temp = str(sh.sheet_arry[sh.sub_sh_count * x_avdd + 2])
                # add the string save sheet name for the usage of plot
                # vout_p_active => is the active sheet (sheet boject)
                # pos_temp is th string of sheet name which used as the reference
                # for plotting function
                # 220825 AVDD and ELVDD is sharing the same sheet assignment since
                # it won't be record at the same time

                if sh.channel_mode == 1:
                    vout_p_pre_active = sh.wb_res.sheets(
                        sh.sheet_arry[sh.sub_sh_count * x_avdd + 3])
                    pos_pre_temp = str(
                        sh.sheet_arry[sh.sub_sh_count * x_avdd + 3])

                elif sh.channel_mode == 0 or sh.channel_mode == 2:
                    vout_p_pre_active = sh.wb_res.sheets(
                        sh.sheet_arry[sh.sub_sh_count * x_avdd + 4])
                    pos_pre_temp = str(
                        sh.sheet_arry[sh.sub_sh_count * x_avdd + 4])

                # 220825 add the sheet mapping for Vout and Von

                if sh.channel_mode == 0 or sh.channel_mode == 2:
                    # only assign the sheet if there are negative output used in the measurement
                    vout_n_active = sh.wb_res.sheets(
                        sh.sheet_arry[sh.sub_sh_count * x_avdd + 3])
                    neg_temp = str(sh.sheet_arry[sh.sub_sh_count * x_avdd + 3])
                    # add the string save sheet name for the usage of plot

                    vout_n_pre_active = sh.wb_res.sheets(
                        sh.sheet_arry[sh.sub_sh_count * x_avdd + 5])
                    neg_pre_temp = str(
                        sh.sheet_arry[sh.sub_sh_count * x_avdd + 5])

                # ====================
                #  this portion seems not a must have portion in the system
                # if sh.channel_mode == 0 or sh.channel_mode == 2 :
                #     # EL only or 3-ch operation
                #     sheet_active = sh.wb_res.sheets(
                #         sh.sheet_arry[sh.sub_sh_count * x_avdd])
                #     raw_active = sh.wb_res.sheets(
                #         sh.sheet_arry[sh.sub_sh_count * x_avdd + 1])
                #     vout_p_active = sh.wb_res.sheets(
                #         sh.sheet_arry[sh.sub_sh_count * x_avdd + 2])
                #     vout_p_active = sh.wb_res.sheets(
                #         sh.sheet_arry[sh.sub_sh_count * x_avdd + 3])
                # else :
                #     # AVDD only
                #     sheet_active = sh.wb_res.sheets(
                #         sh.sheet_arry[sh.sub_sh_count * x_avdd])
                #     raw_active = sh.wb_res.sheets(
                #         sh.sheet_arry[sh.sub_sh_count * x_avdd + 1])
                #     vout_p_active = sh.wb_res.sheets(
                #         sh.sheet_arry[sh.sub_sh_count * x_avdd + 2])
                # ====================

                # not to turn load on here, Vin haven't change yet, only update the
                # AVDD load current here

                # if sim_real == 1 :
                #     load1.chg_out(curr_avdd, sh.loa_ch2_set, 'on')
                #     # turn the load of AVDD on when after load the current
                # else:
                #     print('finished set the current and turn load on')
                #     # input()

                # Vin loop start point
                # counter for the Vin setting

                x_vin = 0
                while x_vin < sh.c_vin:

                    v_target = sh.sh_org_tab.range((3 + x_vin, 2)).value
                    pro_status_str = 'Vin:' + str(v_target)
                    vin_status = str(v_target)
                    program_status(pro_status_str)
                    # update the target Vin to the program status

                    # add the related Vin(ideal) setting at the result sheet
                    sheet_active.range((24, 3 + x_vin)).value = v_target

                    # 220325: regulation sheet also need to have current index
                    if sh.channel_mode == 0 or sh.channel_mode == 2:
                        # ELVDD and ELVSS have regulation sheet
                        vout_p_active.range((24, 3 + x_vin)).value = v_target
                        vout_n_active.range((24, 3 + x_vin)).value = v_target
                        # 220825 add the V setting for Vop and Von sheet
                        vout_p_pre_active.range(
                            (24, 3 + x_vin)).value = v_target
                        vout_n_pre_active.range(
                            (24, 3 + x_vin)).value = v_target
                    else:
                        # AVDD have regulation sheet
                        vout_p_active.range((24, 3 + x_vin)).value = v_target
                        # 220825 add the V setting for Vop and Von sheet
                        vout_p_pre_active.range(
                            (24, 3 + x_vin)).value = v_target

                    # for the raw data sheet index
                    raw_active.range((11 + raw_gap * x_vin, 2)).value = 'Vin'
                    raw_active.range((12 + raw_gap * x_vin, 2)).value = 'Iin'
                    raw_active.range((13 + raw_gap * x_vin, 2)).value = 'ELVDD'
                    raw_active.range((14 + raw_gap * x_vin, 2)).value = 'ELVSS'
                    raw_active.range((15 + raw_gap * x_vin, 2)).value = 'I_EL'
                    raw_active.range((16 + raw_gap * x_vin, 2)).value = 'AVDD'
                    raw_active.range((17 + raw_gap * x_vin, 2)
                                     ).value = 'I_AVDD'
                    raw_active.range((18 + raw_gap * x_vin, 2)).value = 'Eff'
                    raw_active.range((19 + raw_gap * x_vin, 2)).value = 'VOP'
                    raw_active.range((20 + raw_gap * x_vin, 2)).value = 'VON'

                    if sim_real == 1:
                        # adjust the vin voltage
                        pwr1.chg_out(v_target, sh.pre_imax,
                                     sh.inst_pwr_ch3, 'on')
                        time.sleep(wait_small)
                    else:
                        # simulation mode change bias settings
                        print('vin setting change: ' + str(v_target))

                    # current loop start point
                    # counter of current setting
                    x_iload = 0
                    # here means I_EL

                    # selection for current counter based on the channel setting
                    # when testing for AVDD single channel, need to use AVDD current mapping column
                    # for the loop counter
                    c_load_curr = 0
                    if sh.channel_mode == 1:
                        # channel mode: 0-EL, 1-AVDD, 2-EL+AVDD
                        c_load_curr = sh.c_avdd_single
                    else:
                        c_load_curr = sh.c_iload

                    # for each loop of iload, need to define a new offset
                    value_i_offset1 = 0
                    value_i_offset2 = 0
                    while x_iload < c_load_curr:
                        # because different mode need to measure different channel,
                        # rebuild the selection part for items below ... wait for continue (0325)

                        if sh.channel_mode == 0 or sh.channel_mode == 2:
                            # for EL only or 3-ch, load target is for EL power
                            iload_target = sh.sh_org_tab.range(
                                (3 + x_iload, 3)).value
                        else:
                            # if only AVDD efficiency, assign the load target to AVDD_1ch column
                            iload_target = sh.sh_org_tab.range(
                                (3 + x_iload, 7)).value

                        # all i_load target must start from 0 mA for calibration

                        pro_status_str = 'setting iload_target current'
                        i_el_status = str(iload_target)
                        print(pro_status_str)
                        program_status(pro_status_str)

                        if x_vin == 0:
                            sheet_active.range(
                                (25 + x_iload, 2)).value = iload_target
                            # only process once for the current index at the result sheet
                            # 220325: regulation sheet also need to have current index
                            if sh.channel_mode == 0 or sh.channel_mode == 2:
                                # ELVDD and ELVSS have regulation sheet
                                vout_p_active.range(
                                    (25 + x_iload, 2)).value = iload_target
                                vout_n_active.range(
                                    (25 + x_iload, 2)).value = iload_target
                                # 220825 add the Vop and Von
                                vout_p_pre_active.range(
                                    (25 + x_iload, 2)).value = iload_target
                                vout_n_pre_active.range(
                                    (25 + x_iload, 2)).value = iload_target
                            else:
                                # AVDD have regulation sheet
                                vout_p_active.range(
                                    (25 + x_iload, 2)).value = iload_target
                                # 220825 add the Vop and Von
                                vout_p_pre_active.range(
                                    (25 + x_iload, 2)).value = iload_target
                            pro_status_str = 'give index to regulation sheet'
                            print(pro_status_str)
                            program_status(pro_status_str)

                        if sim_mcu == 1:
                            # setting up for the relay channel

                            # uart_cmd_str = (chr(mcu_mode_8_bit_IO) +
                            #                 mcu_cmd_arry[int(meter_ch_ctrl)])
                            # print(uart_cmd_str)
                            # mcu_com.write(uart_cmd_str)
                            # 220328: reset the relay channel to initial state(Vin stage)
                            meter_ch_ctrl = 0
                            mcu_write('relay')
                            time.sleep(wait_small)
                            # input()
                            # finished adjust relay channel

                            # note that the meter_ch_ctrl will change through the measurement channel change
                            # so the the start point of meter should be Vin(calibration)
                        else:
                            print(
                                'change the relay channel without MCU for calibration')
                            # setting up for the relay channel
                            uart_cmd_str = (chr(mcu_mode_8_bit_IO) +
                                            mcu_cmd_arry[int(meter_ch_ctrl)])
                            print(uart_cmd_str)
                            # mcu_com.write(uart_cmd_str)
                            meter_ch_ctrl = 0
                            time.sleep(wait_small)

                        if sim_real == 1:
                            # setting up for the instrument control
                            time.sleep(wait_small)

                            # initialization the temp saving parameter for the efficiency calculation
                            # clear result for each round of the current loop
                            value_elvdd = 0
                            value_elvss = 0
                            value_avdd = 0
                            value_iin = 0
                            value_vin = 0
                            value_iel = 0
                            value_iavdd = 0
                            value_eff = 0
                            if x_iload > 0:
                                # turn the load on to setting i_load (when x > 0)

                                # 20220429 method to assign source meter for the PMIC measurement
                                # only single AVDD efficiency need to use source meter at the AVDD
                                # other is mainly for EL-power
                                # make the selection easier in this applicaation

                                if sh.channel_mode == 2:
                                    # 3 channel mode, only EL powere will have source meter
                                    # selection only at the iload_target but not curr_avdd

                                    # turn on AVDD load channel, must be chroma
                                    load1.chg_out(curr_avdd,
                                                  sh.loa_ch2_set, 'on')

                                    # choose the source meter or the chroma load
                                    if sh.source_meter_channel == 1:
                                        load_src.change_I(iload_target, 'on')
                                        # or you can use 'keep' to replace 'on' and call the load turn at other point

                                    else:
                                        # not using source meter for EL, it's chroma loader
                                        load1.chg_out(
                                            iload_target, sh.loa_ch_set, 'on')

                                elif sh.channel_mode == 0:
                                    # only turn the EL on
                                    # choose the source meter or the chroma load
                                    if sh.source_meter_channel == 1:
                                        load_src.change_I(iload_target, 'on')
                                        # or you can use 'keep' to replace 'on' and call the load turn at other point

                                    else:
                                        # not using source meter for EL, it's chroma loader
                                        load1.chg_out(
                                            iload_target, sh.loa_ch_set, 'on')
                                elif sh.channel_mode == 1:
                                    # only turn the AVDD on
                                    # choose the source meter or the chroma load
                                    if sh.source_meter_channel == 2:
                                        load_src.change_I(iload_target, 'on')
                                        # or you can use 'keep' to replace 'on' and call the load turn at other point

                                    else:
                                        # not using source meter for EL, it's chroma loader
                                        load1.chg_out(
                                            iload_target, sh.loa_ch2_set, 'on')

                            # need to set muc_sim to 1 before using calibration
                            vin_clibrate_singal_met(vin_ch, v_target)
                            # record Vin after calibration finished
                            time.sleep(wait_time)
                            data_latch('vin', v_res_temp)

                            # all the channel change with original sequence
                            # but bypass result to 'NA' with related mode of settings
                            # 0 is bypass and 1 is enable

                            # =====
                            meter_ch_ctrl = meter_ch_ctrl + 1
                            # change relay to AVDD
                            if sh.channel_mode == 3:
                                # bypass AVDD measurement result if only check EL efficiency
                                bypass_measurement_flag = 1
                            else:
                                # only change relay and measurement voltage if needed
                                mcu_write('relay')
                                v_res_temp = met1.mea_v()
                                time.sleep(wait_time)
                            # for the data latch, bypass flag will decide to use the result or not
                            data_latch('avdd', v_res_temp)

                            # =====

                            # =====
                            meter_ch_ctrl = meter_ch_ctrl + 1
                            # change relay to ELVDD
                            if sh.channel_mode == 3:
                                # bypass ELVDD measurement result (AVDD only)
                                bypass_measurement_flag = 1
                            else:
                                # only change relay and measurement voltage if needed
                                mcu_write('relay')
                                v_res_temp = met1.mea_v()
                                time.sleep(wait_time)
                            # for the data latch, bypass flag will decide to use the result or not
                            data_latch('elvdd', v_res_temp)

                            # =====

                            # =====
                            meter_ch_ctrl = meter_ch_ctrl + 1
                            # change relay to ELVSS
                            if sh.channel_mode == 3:
                                # bypass ELVSS measurement result (AVDD only)
                                bypass_measurement_flag = 1
                            else:
                                # only change relay and measurement voltage if needed
                                mcu_write('relay')
                                v_res_temp = met1.mea_v()
                                time.sleep(wait_time)
                            # for the data latch, bypass flag will decide to use the result or not
                            data_latch('elvss', v_res_temp)

                            # =====

                            # =====
                            meter_ch_ctrl = meter_ch_ctrl + 1
                            # change relay to VOP
                            if sh.channel_mode == 3:
                                # record in all condition
                                bypass_measurement_flag = 1
                            else:
                                # only change relay and measurement voltage if needed
                                mcu_write('relay')
                                v_res_temp = met1.mea_v()
                                time.sleep(wait_time)
                            # for the data latch, bypass flag will decide to use the result or not
                            data_latch('vop', v_res_temp)

                            # =====

                            # =====
                            meter_ch_ctrl = meter_ch_ctrl + 1
                            # change relay to VON
                            if sh.channel_mode == 3:
                                # bypass VON measurement result (AVDD only)
                                bypass_measurement_flag = 1
                            else:
                                # only change relay and measurement voltage if needed
                                mcu_write('relay')
                                v_res_temp = met1.mea_v()
                                time.sleep(wait_time)
                            # for the data latch, bypass flag will decide to use the result or not
                            data_latch('von', v_res_temp)

                            # =====

                            # # this is the power supply method, not the meter method
                            # v_res_temp = pwr1.read_iout()
                            # # for current reading, need to remove A in the end of string
                            # v_res_temp = v_res_temp.replace('A', '')
                            # # this part can also consider to move to the next ints_pkg file
                            # # can help to improve the complexity

                            # adjust the Iin measurement from power supply to meter
                            v_res_temp = met2.mea_i()
                            data_latch('iin', v_res_temp)
                            time.sleep(wait_time)
                            # different mode need different Iout => read all Iout but only keep the good one

                            # 20220429 read I function: read all the channel, but choose to latch or not,
                            # get the I read result from both chroma channel, choose different way to latch data
                            # based on the channel mode from control sheet
                            if sh.source_meter_channel == 0:
                                v_res_temp = load1.read_iout(sh.loa_ch_set)
                                if sh.channel_mode == 0 or sh.channel_mode == 2:
                                    data_latch('i_el', v_res_temp)
                                    if sh.channel_mode == 0:
                                        # give the i_avdd blank to 0 for result
                                        data_latch(
                                            'i_avdd', str(value_i_offset2))
                                        # 0511 to preven calibration settings cause error
                                        # pass the calibration parameter into data_latch to
                                        # cancel the result adjustment
                                v_res_temp = load1.read_iout(sh.loa_ch2_set)
                                if sh.channel_mode == 1 or sh.channel_mode == 2:
                                    data_latch('i_avdd', v_res_temp)
                                    if sh.channel_mode == 1:
                                        # give the i_el blank to 0 for result
                                        data_latch('i_el', '0')
                                        data_latch(
                                            'i_el', str(value_i_offset1))
                                        # 0511 to preven calibration settings cause error
                                        # pass the calibration parameter into data_latch to
                                        # cancel the result adjustment
                            elif sh.source_meter_channel == 1:
                                v_res_temp = load_src.read('CURR')
                                data_latch('i_el', v_res_temp)
                                if sh.channel_mode == 2 or sh.channel_mode == 0:
                                    # if now is 3-channel mode, also need to latch the current at AVDD
                                    v_res_temp = load1.read_iout(
                                        sh.loa_ch2_set)
                                    data_latch('i_avdd', v_res_temp)
                            elif sh.source_meter_channel == 2:
                                # since we already assume there is no EL power measurement
                                # for using source meter for AVDD, EL power will be latch to 0
                                v_res_temp = load_src.read('CURR')
                                data_latch('i_avdd', v_res_temp)
                                # give the i_el blank to 0 for result
                                data_latch('i_el', '0')

                            # ==== loader offset configuration
                            # add the calibration factor to iload_target
                            if x_iload == 0:

                                if sh.loader_cal_mode == 2:
                                    # when the calibration mode is 2, use no load case for the offset adjustment
                                    if sh.channel_mode == 0 or sh.channel_mode == 2:
                                        value_i_offset1 = value_iel
                                        if sh.channel_mode == 2:
                                            value_i_offset2 = value_iavdd
                                    elif sh.channel_mode == 1:
                                        value_i_offset2 = value_iavdd
                                elif sh.loader_cal_mode == 1:
                                    # when the calibration mode is 1, constant offset adjustment
                                    if sh.channel_mode == 0 or sh.channel_mode == 2:
                                        value_i_offset1 = sh.loader_cal_off1
                                        if sh.channel_mode == 2:
                                            value_i_offset2 = sh.loader_cal_off2
                                    elif sh.channel_mode == 1:
                                        value_i_offset2 = sh.loader_cal_off2
                                else:
                                    # the case don't need offset
                                    value_i_offset1 = 0
                                    value_i_offset2 = 0

                                if sh.source_meter_channel == 1:
                                    # if the channel 1=> EL power, 2=> AVDD, 0=> not use source meter
                                    value_i_offset1 = 0
                                    # remove the offset if using source meter
                                elif sh.source_meter_channel == 2:
                                    value_i_offset2 = 0

                            # ==== loader offset end point
                            print(sh.sh_main)
                            pass

                            # 220329: for the new version of inst_pkg_a, add the A remove function in
                            # the sub-program operation
                            # # for current reading, need to remove A in the end of string
                            # v_res_temp = v_res_temp.replace('A', '')
                            # data_latch('iout', v_res_temp)
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

                            # to prevent overspec of the votlage input, need to change the Vin back to
                            # intital target before loading release
                            pwr1.chg_out(v_target, sh.pre_sup_iout,
                                         sh.inst_pwr_ch3, 'on')
                            print('V_target return to normal')

                            # release loading
                            # turn the load off after measurement
                            if sh.channel_mode == 2:
                                # load1.chg_out(curr_avdd, sh.loa_ch2_set, 'off')
                                # load1.chg_out(iload_target, sh.loa_ch_set, 'off')
                                load1.chg_out(0, sh.loa_ch2_set, 'on')
                                load1.chg_out(0, sh.loa_ch_set, 'on')
                            elif sh.channel_mode == 0:
                                # only turn the EL on
                                # load1.chg_out(iload_target, sh.loa_ch_set, 'off')
                                if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
                                    # load_src.load_off()
                                    # change to turn off at each voltage cycle for loadr and source meter
                                    load_src.change_I(0, 'on')
                                else:
                                    load1.chg_out(0, sh.loa_ch_set, 'on')
                            elif sh.channel_mode == 1:
                                # only turn the AVDD on
                                # load1.chg_out(iload_target, sh.loa_ch2_set, 'off')
                                if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
                                    # load_src.load_off()
                                    # change to turn off at each voltage cycle for loadr and source meter
                                    load_src.change_I(0, 'on')
                                else:
                                    load1.chg_out(0, sh.loa_ch2_set, 'on')

                            # 20220429 since release the load and set to turn off is ok,
                            # no specific setting for the chroma load selection here
                            # source meter is also turn off directly

                            # 220824 to prevent the wrong turning off of the E-load when using source meter for sngle channel operation
                            # need to change this source meter turn off into loader turn off
                            # if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
                            #     # load_src.load_off()
                            #     # change to turn off at each voltage cycle for loadr and source meter
                            #     load_src.change_I(0, 'on')

                            # after the result fix in the data saving excel, calculate the efficiency
                            value_eff = ((value_elvdd - value_elvss) * value_iel +
                                         value_avdd * value_iavdd) / (value_vin * value_iin)
                            data_latch('eff', str(value_eff))
                            # latch and locked down the efficiency result

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
                                # initial by the previous code
                                mcu_write('relay')
                                print('change relay, meter channel change to: ' +
                                      str(meter_ch_ctrl))
                                print('now is Vin')

                                # testing for the MCU operation independently
                                meter_ch_ctrl = meter_ch_ctrl + 1
                                # change relay to Vbias
                                mcu_write('relay')
                                print('change relay, meter channel change to: ' +
                                      str(meter_ch_ctrl))
                                print('now is AVDD')
                                # input()

                                # testing for the MCU operation independently
                                meter_ch_ctrl = meter_ch_ctrl + 1
                                # change relay to Vbias
                                mcu_write('relay')
                                print('change relay, meter channel change to: ' +
                                      str(meter_ch_ctrl))
                                print('now is OVDD')
                                # input()

                                # testing for the MCU operation independently
                                meter_ch_ctrl = meter_ch_ctrl + 1
                                # change relay to Vbias
                                mcu_write('relay')
                                print('change relay, meter channel change to: ' +
                                      str(meter_ch_ctrl))
                                print('now is OVSS')
                                # input()

                                # testing for the MCU operation independently
                                meter_ch_ctrl = meter_ch_ctrl + 1
                                # change relay to Vbias
                                mcu_write('relay')
                                print('change relay, meter channel change to: ' +
                                      str(meter_ch_ctrl))
                                print('now is VOP')
                                # input()

                                # testing for the MCU operation independently
                                meter_ch_ctrl = meter_ch_ctrl + 1
                                # change relay to Vbias
                                mcu_write('relay')
                                print('change relay, meter channel change to: ' +
                                      str(meter_ch_ctrl))
                                print('now is VON')
                                # input()

                            data_latch('vin', str(sim_v_data_temp))
                            v_res_temp = v_res_temp + 1
                            sim_v_data_temp = v_res_temp + 0.01
                            data_latch('iin', str(sim_v_data_temp))
                            if sh.channel_mode == 0 or sh.channel_mode == 2:
                                # data_latch only active when using AVDD mode(1) or 3-ch mode(2)
                                v_res_temp = v_res_temp + 1
                                sim_v_data_temp = v_res_temp + 0.02
                                data_latch('elvdd', str(sim_v_data_temp))
                                v_res_temp = v_res_temp + 1
                                sim_v_data_temp = v_res_temp + iload_target
                                data_latch('elvss', str(sim_v_data_temp))
                            v_res_temp = v_res_temp + 1
                            sim_v_data_temp = v_res_temp + 0.03
                            if sh.source_meter_channel == 1:
                                print('i_el is from the source meter')
                                print('')
                                pass
                            data_latch('i_el', str(sim_v_data_temp))
                            if sh.channel_mode == 1 or sh.channel_mode == 2:
                                # data_latch only active when using AVDD mode(1) or 3-ch mode(2)
                                v_res_temp = v_res_temp + 1
                                sim_v_data_temp = v_res_temp + 0.04
                                data_latch('avdd', str(sim_v_data_temp))
                            v_res_temp = v_res_temp + 1
                            sim_v_data_temp = v_res_temp + 0.05
                            if sh.source_meter_channel == 2 and sh.channel_mode == 1:
                                print('i_avdd is from the source meter')
                                print('')
                                pass
                            data_latch('i_avdd', str(sim_v_data_temp))
                            v_res_temp = v_res_temp + 1
                            sim_v_data_temp = v_res_temp + 0.06
                            data_latch('eff', str(sim_v_data_temp))
                            v_res_temp = v_res_temp + 1
                            sim_v_data_temp = v_res_temp + 0.07
                            data_latch('vop', str(sim_v_data_temp))
                            v_res_temp = v_res_temp + 1
                            if sh.channel_mode == 0 or sh.channel_mode == 2:
                                sim_v_data_temp = v_res_temp + 0.08
                                data_latch('von', str(sim_v_data_temp))

                            # the simulation mode for result generation
                            print(sh.sh_main)
                            print('')

                        # before x_iload increase, need to checck inst status based on insturment control
                        # 220522 add the check inst sub program
                        if sh.inst_auto_selection == 2:
                            check_inst_update()
                        if sh.program_exit == 0:
                            # exit the program
                            break
                        # check every time goes to the loop

                        x_iload = x_iload + 1
                        # end of the 4th loop

                    sh.wb_res.save(sh.result_book_trace)
                    if sh.program_exit == 0:
                        # exit the program
                        break
                    x_vin = x_vin + 1
                    # end of the 3rd loop

                # sh.wb_res.save(sh.result_book_trace)
                # save the file one time after each avdd load current is finished

                # make the plot for each AVDD current change, since there are one group of the efficiency and regulation data
                # start from EL only but need to check all 3 mode of the operation
                v_cnt = sh.c_vin
                i_cnt = c_load_curr

                # decide from the sheet need to plot chart
                book_n = str(sh.full_result_name) + '.xlsx'

                # 220524: for the instrument operation, prevent the issue of change windoww,
                # need to have message remind user to release control of excel
                # so there will not be the error from the plot
                # remind the operation can keep going after the plot is finished
                if sh.en_plot_waring == 1:
                    msg_res = win32api.MessageBox(
                        0, 'Release the control of excel, change the window to auto file now, and press enter, remind again when the plot is finished', 'Plot request from python')

                # plot for efficiency
                sheet_n = eff_temp
                sh.excel.Application.Run("EFF_inst.xlsm!gary_chart",
                                         v_cnt, i_cnt, sheet_n, book_n, y_raw_start, x_raw_start)

                if sh.channel_mode == 0 or sh.channel_mode == 2:

                    # plot for ELVDD
                    sheet_n = pos_temp
                    sh.excel.Application.Run("EFF_inst.xlsm!gary_chart",
                                             v_cnt, i_cnt, sheet_n, book_n, y_raw_start, x_raw_start)

                    # plot for ELVSS
                    sheet_n = neg_temp
                    sh.excel.Application.Run("EFF_inst.xlsm!gary_chart",
                                             v_cnt, i_cnt, sheet_n, book_n, y_raw_start, x_raw_start)

                    # plot for Vout
                    sheet_n = pos_pre_temp
                    sh.excel.Application.Run("EFF_inst.xlsm!gary_chart",
                                             v_cnt, i_cnt, sheet_n, book_n, y_raw_start, x_raw_start)

                    # plot for Von
                    sheet_n = neg_pre_temp
                    sh.excel.Application.Run("EFF_inst.xlsm!gary_chart",
                                             v_cnt, i_cnt, sheet_n, book_n, y_raw_start, x_raw_start)
                else:

                    # plot for AVDD
                    sheet_n = pos_temp
                    sh.excel.Application.Run("EFF_inst.xlsm!gary_chart",
                                             v_cnt, i_cnt, sheet_n, book_n, y_raw_start, x_raw_start)

                    # plot for Vout
                    sheet_n = pos_pre_temp
                    sh.excel.Application.Run("EFF_inst.xlsm!gary_chart",
                                             v_cnt, i_cnt, sheet_n, book_n, y_raw_start, x_raw_start)

                sh.wb_res.save(sh.result_book_trace)
                print(sh.sh_main)
                print('')
                # save the result after plot is finished

                # the plot request is finished and jump another window to remind
                if sh.en_plot_waring == 1:
                    msg_res = win32api.MessageBox(
                        0, 'You can start to operate the computer again', 'Plot request finished ')

                if sh.program_exit == 0:
                    # exit the program
                    break
                x_avdd = x_avdd + 1
                # end of the 2nd loop
            if sh.program_exit == 0:
                # exit the program
                break

            x_sw_i2c = x_sw_i2c + 1
            # end of the 1st loop

        # turn off the load and source after the loop is finished
        if sim_real == 1:

            # turn off the power and load
            pwr1.chg_out(0, sh.pre_imax, sh.inst_pwr_ch3, 'off')
            load1.chg_out(0, sh.loa_ch_set, 'off')
            load1.chg_out(0, sh.loa_ch2_set, 'off')

            if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
                load_src.load_off()

        eff_done = 1
        # eff done is set after one cycle of efficiency measurement
        # efficiency measurement will need to reset
        # change the eff_done status on the main sheet
        sh.eff_done_sh = 1
        sh.sh_main.range('C13').value = 1
        print(sh.sh_main)
        print('')

        # 220824 add the exit action for save and turn off the excel
        if sh.sheet_off_finished == 1:
            sh.wb_res.save(sh.result_book_trace)
            sh.wb_res.close()
        else:
            pass
        # to preven the issue of re-run (same file opening and crash)


# simulation mode error => need to add the
if sim_real == 1:
    # turn off all the instrument
    pwr1.inst_close()
    load1.inst_close()
    if sh.source_meter_channel == 1 or sh.source_meter_channel == 2:
        load_src.inst_close()

else:
    print('now is going to turn off the instrument')
    print('this is the end of simulation mode ')

if sh.en_fully_auto == 1:
    print('fully auto mode enable~going to turn off')
    # close the control book when testing finished
    sh.wb.clsoe()
    time.sleep(3)


print('finsihed and goodbye')
