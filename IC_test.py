# IC_test :
# tthis program used to check if IC is borken after testing


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
import sheet_ctrl_ICtest as sh
# loading of information of contorol book and parameter
import inst_pkg as inst
# loading for instrument object ()
# controlled by simulation mode variable

import locale as lo
# include for atof function => transfer string to float

wait_time = sh.wait_time
wait_small = 0.2
wait_sim = 0.02  # waiting time for simulation mode test
# general waiting time for slppe function

v_res_temp = 0
# measurement result temp variable


sim_real = 1
# 1 => real mode, 0=> simulation(program debug) mode
sim_mcu = 1
# this is only used for MCU com port test simuation
# when setting (sim_real, sim_mcu) = (0, 1)
# enable MCU testing when GPIC disable
# because NCU will be separate with GPIB for implementation and test
# mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']

# 220528 : update the new sequence based on the fixed relay mapping
mcu_cmd_arry = ['02', '04', '08', '10', '20', '01', '40', '80']

meter_ch_ctrl = 0
# meter channel indicator: 0: AVDD, 1: ELVDD, 2: ELVSS, 3: VOUT, 4: VON
# 220528: related setup for general setting:
# 0 change to Vin, other sequence are the same
# 0 : Vin, 1: AVDD, 2: ELVDD, 3: ELVSS, 4: VOUT, 5: VON, 6: PWR_CH1, 7: PWR_CH2

mode_iq = 3
mode_ra = 5
mode_sw = 1
# MCU mapping for RA_GPIO control is mode 2 (SWIRE scan)

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

    # MCU initial for IQ test under normal mode
    uart_cmd_str = chr(mode_iq) + chr(4) + chr(1)
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)
    time.sleep(wait_time)
    # input()

# MCU need to cover the change from positive to negative
# ELVDD, ELVSS and AVDD channel change
# MCU channel, initial setting:


# real mode using instrument connection
if sim_real == 1:
    # import inst_pkg as inst

    pwr1 = inst.LPS_505N(0, 0, 1, sh.pwr_sup_addr, 'off')
    # no worries about the channel of supply, will change at initial state below

    met1 = inst.Met_34460(sh.met_vin_res, sh.met_vin_rang,
                          sh.met_iin_res, sh.met_iin_rang, sh.met_vdd_addr)
    # meter loaded from the main sheet
    met2 = inst.Met_34460(sh.met_vin_res, sh.met_vin_rang,
                          sh.met_iin_res, sh.met_iin_rang, sh.met_vss_addr)
    if sh.loa_en == 1:
        # loader is enable in the testing
        load1 = inst.chroma_63600(
            1, sh.loa_dts_addr, str(sh.loa_mod_set))
    # # load channel and related setting also change after definition

    # turn on the GPIB connection of instrument
    pwr1.open_inst()
    met1.open_instr()
    if sh.loa_en == 1:
        # 20211207 update: for the correction of open inst function for loader
        load1.open_inst()
    met2.open_instr()

    # # load the name to related blank in result book
    # sh.sh_main.range('D27').value = pwr1.inst_name()
    # sh.sh_main.range('D28').value = met1.inst_name()
    # sh.sh_main.range('D30').value = load1.inst_name()

    # the other way for update instrument name(sub program in sh_ctrl)
    sh.inst_name_sheet('PWR1', pwr1.inst_name())
    sh.inst_name_sheet('MET1', met1.inst_name())
    sh.inst_name_sheet('MET2', met2.inst_name())
    if sh.loa_en == 1:
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

    # should not need the extra Vin in the change
    # pwr1.chg_out(sh.vin1_set, sh.pre_sup_iout, sh.pwr_ch_set, 'on')

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
        # pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.pwr_ch_set, 'on')
        print('pre-power on here')

        msg_res = win32api.MessageBox(0, 'press enter if hardware configuration is correct',
                                      'Pre-power on for system test under Vin= ' + str(sh.pre_vin) + 'Iin= ' + str(sh.pre_sup_iout))

    print('pre-power on state finished and ready for next')
    # can used any inpuy for pore-power on finished test ing on excel
    sh.sh_org_tab.range((4, 4)).value = 'test mode'
    sh.sh_org_tab.range((5, 4)).value = 'test ok'
    # RA sheet also input the testing information for reference
    sh.sh_ra_tab.range((9, 1)).value = 'testing finished'

    # input()


# after pre-power on, initial the setting of load channel
# 2- channel for EL power and AVDD


# the end of instrument initialization, alread after pre-power on

# no dynamic board needed, just follow the sequence
if sim_real == 1:
    # config the load for EL power and AVDD if load enable
    if sh.loa_en == 1:
        # only set up the current here, not on yet
        load1.chg_out(sh.loa_curr, sh.loa_ch_set, 'off')
        load1.chg_out(sh.loa_curr2, sh.loa_ch2_set, 'off')

    # turn the power supply on
    pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.pwr_ch_set, 'on')
    # delay after power on to prevent transient measurement
    time.sleep(wait_time)

else:
    print('the initialization of power supplty and loader is finished')
    # input()

# RA test settings before the loop
# items counter (from RA_output, total 6, 1 current + 5 voltage)
RA_out_V = 6
v_sim = 0

if sh.before_after == 0:
    y_start = 4
    x_start = 3
elif sh.before_after == 1:
    y_start = 4
    x_start = 11

c_board = 0
while c_board < sh.total_board:

    sh.sh_org_tab.range('C3').value = 'waitting'
    sh.sh_ra_tab.range('A8').value = 'waitting'
    if sim_real == 1:
        msg_res = win32api.MessageBox(
            0, 'press enter after new board is ready', 'Change to the next board')
    else:
        time.sleep(wait_sim)

    sh.sh_ra_tab.range('A8').value = 'busy'
    sh.sh_org_tab.range('C3').value = 'busy'
    # need to create trigger at each board
    if sh.single_test != 3:
        # =============== IQ test
        # start for the IQ measured part, copy from the IQ_test

        x_iq = 0
        # counter for SWIRE pulse amount
        while x_iq < sh.c_iq:
            # load the Vin command first

            # shift of loading table is following the different board, consider counter and y range change)
            ideal_v = sh.sh_org_tab.range(
                (7 + c_board * (sh.tab_cp_y_ran + sh.tab_cp_dist), 4 + x_iq)).value

            if sim_real == 1:
                # update the vin setting for dif)ferent vin demand
                pwr1.chg_out(ideal_v, sh.pre_sup_iout, sh.pwr_ch_set, 'on')

            else:
                # simulation mode update the Vin mapping result
                sh.sh_org_tab.range(
                    (6 + c_board * (sh.tab_cp_y_ran + sh.tab_cp_dist), 4 + x_iq)).value = ideal_v
            # four different mode in this loop will change

            x_submode = 0
            # initialize the counter for differnt state
            while x_submode < 4:
                # submode + 1 will be the state command for MCU

                if sim_mcu == 1:

                    uart_cmd_str = (chr(int(mode_iq)) +
                                    chr(x_submode + 1) + chr(1))
                    # for IQ scan only need one state command to MCU
                    print(uart_cmd_str)
                    mcu_com.write(uart_cmd_str)
                    time.sleep(wait_time)
                    if x_submode == 0:
                        time.sleep(wait_time)
                        # extra wait time for the normal mode to shtudown mode transition
                        # because IQ may not stop change so fast, need to double check

                else:
                    # items check in simulation mode
                    uart_cmd_str = (chr(int(mode_iq)) +
                                    chr(x_submode + 1) + chr(1))
                    print(uart_cmd_str)

                    print('input voltage setting is ' + str(ideal_v))
                    print('the mode - 1 is ' + str(x_submode))
                    time.sleep(wait_sim)
                    # input()

                if sim_real == 1:

                    # this part can not be in the simulation mode,
                    # because it need to access the meter object variable
                    if x_submode == 0:
                        # for the mode of measure ISD, need to have more wait time for stable and
                        # need to prevend negative result of error result

                        # here used the change of range directly => change variable in inst_pkg
                        range_temp = met1.max_mea_i_ini
                        met1.max_mea_i_ini = sh.iq_range

                        v_res_temp = met1.mea_i()
                        # call measurement to update the ISD measurement range on the meter

                        if sh.ISD_hand == 1:
                            msg_res = win32api.MessageBox(
                                0, 'press enter after finished key in ISD', 'SA record ISD in hand')
                            # jump the message out for IQ record
                            pass
                        else:
                            time.sleep(5 * wait_time)
                            # give some to stable for ISD measurement
                            pass

                    else:
                        met1.max_mea_i_ini = range_temp
                        # because the counter starts from 0, ini value will save to the temp first and return when the counter
                        # is not 0

                    # measurement start after the AVDDEN and SWIRE is updated
                    time.sleep(wait_small)
                    v_res_temp = met1.mea_i()

                    if x_submode == 0 and sh.ISD_hand == 0:
                        while float(v_res_temp) < 0:
                            # when there is a negative result from the measurement
                            # we need to re measure to update the result
                            v_res_temp = met1.mea_i()
                            time.sleep(wait_small)
                            # it should be already stable after the change of GPIO command
                            # small wait time should be enough

                    # when the measurement is finished, update the result to excel table and map to the scaling
                    time.sleep(wait_small)
                    if x_submode > 0:
                        # other mode rather then ISD, record directly
                        sh.sh_org_tab.range((8 + c_board * (sh.tab_cp_y_ran + sh.tab_cp_dist) + x_submode, 4 + x_iq)
                                            ).value = lo.atof(v_res_temp) * sh.iq_scaling
                        # iq_scaling is decide from the result table (unit is optional)
                        pass
                    else:
                        # case when submode is 0 (ISD)
                        if sh.ISD_hand == 1:
                            pass
                        elif sh.ISD_hand == 0:
                            # need to record the ISD measured by the meter
                            sh.sh_org_tab.range((8 + c_board * (sh.tab_cp_y_ran + sh.tab_cp_dist) +
                                                x_submode, 4 + x_iq)).value = lo.atof(v_res_temp) * sh.iq_scaling
                            pass
                        pass

                else:
                    # finished the measurement at simulation mode
                    print('finished the simulation mode at measurement')
                    sh.sh_org_tab.range((8 + x_submode + c_board * (sh.tab_cp_y_ran + sh.tab_cp_dist), 4 + x_iq)
                                        ).value = x_submode + 1 + 10 * x_iq
                    print(sh.sh_org_tab.range((8 + x_submode + c_board * (sh.tab_cp_y_ran + sh.tab_cp_dist), 4 + x_iq)
                                              ).value)
                    # simulation mode result

                x_submode = x_submode + 1

            # update the counter for different Vin

            if sim_real == 0:
                # simulation mode update the Vin mapping result
                sh.sh_org_tab.range(
                    (6 + c_board * (sh.tab_cp_y_ran + sh.tab_cp_dist), 4 + x_iq)).value = ideal_v

            x_iq = x_iq + 1
            # save the result after each counter finished
            # time.sleep(wait_sim)
            # sh.sh_ra_tab.activate
            # sh.wb_res.save(sh.result_book_trace)
            # sh.wb_res.save()

        # =============== IQ test end

    if sh.single_test != 2:
        # =============== RA test

        # the the control loop of board should move to the top of IQ measurement
        # c_board = 0
        # while c_board < sh.total_board:

        # setup Vin but no calibration => real operation
        # assign board number
        sh.sh_ra_tab.range((y_start + int(c_board),
                            x_start - 1)).value = c_board + 1

        # setup the power supply and COM port SWIRE command at every board operation

        # power supply configuration and turn on
        if sim_real == 1:
            pwr1.chg_out(sh.pre_vin, sh.pre_sup_iout, sh.pwr_ch_set, 'on')
            # delay after power on to prevent transient measurement
            time.sleep(wait_time)
            pwr1.change_V(sh.Vin)
            time.sleep(wait_time)
            # add the loading if needed
            if sh.loa_en == 1:
                load1.chg_out(sh.loa_curr, sh.loa_ch_set, 'on')
                load1.chg_out(sh.loa_curr2, sh.loa_ch2_set, 'on')

        if sim_mcu == 1:
            # SWIRE command for the maximum output voltage of ELVDD and ELVSS
            # SWIRE default status need to be high
            uart_cmd_str = chr(int(mode_sw)) + chr(int(sh.pulse_ELVSS)) + \
                chr(int(sh.pulse_ELVDD))
            print(uart_cmd_str)
            mcu_com.write(uart_cmd_str)
            time.sleep(wait_time)
            if sh.pulse_AVDD != 0:
                uart_cmd_str = chr(int(mode_sw)) + chr(int(sh.pulse_AVDD)) + \
                    chr(int(sh.pulse_AVDD))
                print(uart_cmd_str)
                mcu_com.write(uart_cmd_str)
                time.sleep(wait_time)

            # input()
        c_item = 0
        # reset counter before each board
        while c_item < RA_out_V:

            # command sequence should map with the blank in excel sheet to output

            if c_item == 0 and sim_real == 1:
                # when c_item = 0, read the current from power supply
                # read the curent and put to related blank in the sheet
                v_sim = pwr1.read_iout()
                v_sim = v_sim.replace('A', '')
                time.sleep(wait_time)
                # right shift 1 block for the
                # sh_org_tab.range((y_start + int(c_dy_board) - 1, x_start + c_item)).value =
                # no need to update data in the selection, but before the end of loop
                # elif sim_real == 0 :
                # no need else if and change the update v_sim at the end of loop

            elif c_item != 0:
                # adjust the output channel first
                if sim_mcu == 1:
                    # first c_item is reading iout
                    mcu_com.write(chr(int(mode_ra)) + mcu_cmd_arry[c_item - 1])
                    print(chr(int(mode_ra)) + mcu_cmd_arry[c_item - 1])
                    time.sleep(wait_time)
                    # input()
                # when c_item != 0, read the voltage from related meter
                if sim_real == 1:
                    if sh.single_test == 1:
                        msg_res = win32api.MessageBox(
                            0, 'press enter when change test point', 'record the next voltage')
                    v_sim = met2.mea_v()
                    time.sleep(wait_time)
                # # MCU UART output only for the output votlage measurement
                # if sim_mcu == 1:
                #     mcu_com.write(str(c_item))
                #     print('mcu out:'+str((c_item)))
                #     time.sleep(wait_time)

            if sim_real == 0:
                # update result in simulation mode
                v_sim = v_sim + 1
                if sh.single_test == 1:
                    msg_res = win32api.MessageBox(
                        0, 'press enter when change test point', 'record the next voltage')
                sh.sh_ra_tab.range((y_start + int(c_board),
                                    x_start + c_item)).value = v_sim
                time.sleep(wait_sim)
            if sim_real == 1:
                # update the data to sheet here
                sh.sh_ra_tab.range(
                    (y_start + c_board, x_start + c_item)).value = lo.atof(v_sim)
            # real data won't need to change v_sim by count
            # v_sim = v_sim + 1
            c_item = c_item + 1
            # update counter at the end of loop

        # power supply need to turn off here
        # for changing the next board
        if sim_real == 1:
            pwr1.change_V(0)
            # turn off the load if load is used in this case
            if sh.loa_en == 1:
                load1.chg_out(sh.loa_curr, sh.loa_ch_set, 'off')
                load1.chg_out(sh.loa_curr2, sh.loa_ch2_set, 'off')

    c_board = c_board + 1
    # time.sleep(wait_time)

    # update counter at the end of loop

    # if the load_en is high, we will use the full load
    # condition to test on the voltage

    # for the RA part, need to think about if going to follow
    # the original table, because there are only one board under testing?
    # or if there are many board?
    # then the IQ table need to change


sh.wb_res.save()
# remember to save the file at the end(no path means current path)
sh.wb_res.close()
