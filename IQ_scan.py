# IQ_scan:
# this program change AVDDEN and SWIRE to check diferent Iin at each Vin and different mode

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
import sheet_ctrl_IQ as sh
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


sim_real = 1
# 1 => real mode, 0=> simulation(program debug) mode
sim_mcu = 1
# this is only used for MCU com port test simuation
# when setting (sim_real, sim_mcu) = (0, 1)
# enable MCU testing when GPIC disable
# because NCU will be separate with GPIB for implementation and test
mcu_cmd_arry = ['01', '02', '04', '08', '10', '20', '40', '80']

meter_ch_ctrl = 0
# meter channel indicator: 0: AVDD, 1: ELVDD, 2: ELVSS

mode = 3
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
    uart_cmd_str = chr(mode) + chr(4) + chr(1)
    print(uart_cmd_str)
    mcu_com.write(uart_cmd_str)
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

    # load1 = inst.chroma_63600(
    #     1, sh.loa_dts_addr, str(sh.loa_mod_set))
    # # load channel and related setting also change after definition

    # turn on the GPIB connection of instrument
    pwr1.open_inst()
    met1.open_instr()
    # load1.open_inst()

    # # load the name to related blank in result book
    # sh.sh_main.range('D27').value = pwr1.inst_name()
    # sh.sh_main.range('D28').value = met1.inst_name()
    # sh.sh_main.range('D30').value = load1.inst_name()

    # the other way for update instrument name(sub program in sh_ctrl)
    sh.inst_name_sheet('PWR1', pwr1.inst_name())
    sh.inst_name_sheet('MET1', met1.inst_name())
    # sh.inst_name_sheet('LOAD1', load1.inst_name())

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
    # input()

x_iq = 0
# counter for SWIRE pulse amount
while x_iq < sh.c_iq:
    # load the Vin command first
    ideal_v = sh.sh_org_tab.range((7, 4 + x_iq)).value

    if sim_real == 1:
        # update the vin setting for different vin demand
        pwr1.chg_out(ideal_v, sh.pre_sup_iout, sh.pwr_ch_set, 'on')

    else:
        # simulation mode update the Vin mapping result
        sh.sh_org_tab.range((6, 4 + x_iq)).value = ideal_v
    # four different mode in this loop will change

    x_submode = 0
    # initialize the counter for differnt state
    while x_submode < 4:
        # submode + 1 will be the state command for MCU

        if sim_mcu == 1:

            uart_cmd_str = (chr(mode) + chr(x_submode + 1) + chr(1))
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
            uart_cmd_str = (chr(mode) + chr(x_submode + 1) + chr(1))
            print(uart_cmd_str)

            print('input voltage setting is ' + str(ideal_v))
            print('the mode - 1 is ' + str(x_submode))
            time.sleep(wait_time)
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
                time.sleep(5 * wait_time)
            else:
                met1.max_mea_i_ini = range_temp
                # because the counter starts from 0, ini value will save to the temp first and return when the counter
                # is not 0

            # measurement start after the AVDDEN and SWIRE is updated
            time.sleep(wait_small)
            v_res_temp = met1.mea_i()

            if x_submode == 0:
                while float(v_res_temp) < 0:
                    # when there is a negative result from the measurement
                    # we need to re measure to update the result
                    v_res_temp = met1.mea_i()
                    time.sleep(wait_small)
                    # it should be already stable after the change of GPIO command
                    # small wait time should be enough

            # when the measurement is finished, update the result to excel table and map to the scaling
            time.sleep(wait_small)
            sh.sh_org_tab.range((8 + x_submode, 4 + x_iq)
                                ).value = lo.atof(v_res_temp) * sh.iq_scaling
            # iq_scaling is decide from the result table (unit is optional)

        else:
            # finished the measurement at simulation mode
            print('finished the simulation mode at measurement')
            sh.sh_org_tab.range((8 + x_submode, 4 + x_iq)
                                ).value = x_submode + 1 + 10 * x_iq
            print(sh.sh_org_tab.range((8 + x_submode, 4 + x_iq)
                                      ).value)
            # simulation mode result

        x_submode = x_submode + 1

    # update the counter for different Vin

    if sim_real == 0:
        # simulation mode update the Vin mapping result
        sh.sh_org_tab.range((6, 4 + x_iq)).value = ideal_v

    x_iq = x_iq + 1
    # save the result after each counter finished
    sh.wb_res.save(sh.result_book_trace)

    # add the command to turn off the power supply after test is finished
    if sim_real == 1:
        pwr1.chg_out(0, sh.pre_sup_iout, sh.pwr_ch_set, 'off')
        # pwr1.inst_close()
        # since inst_close may turn all the channel, may not be a good command for single function
print('finsihed and goodbye')
