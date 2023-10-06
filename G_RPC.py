from multiprocessing.connection import Client, wait


# excel parameter and settings
import parameter_load_obj as par
import time
import xlwings as xw

# === other support object
# add the report_arragement for building the result in the report sheet
import report_arragement_obj as rep_obj
import JIGM3 as mcu_g


# === initialization for the support object
file_name = "c:\\py_gary\\test_excel\\GPL_V5_RPC_temp.xlsx"
rep_a = rep_obj.report_arragement(file_name0=file_name, full_trace0=1)

mcu_m = mcu_g.JIGM3(sim_mcu0=0)
print("JIGM3 MCU selected for Grace")

NAGSIGN_RPC = "##NAGRPC##"
# fmt: off

class NAGuiRPC:
    def __init__(self, timeout=3.0, excel0=0, pwr0=0, met_v0=0, loader_0=0, mcu0=0, src0=0, met_i0=0, chamber0=0, scope0=0):
        '''
        input all the instrument may need in RPC application
        231002: add Gary's testing object initial items
        '''

        self.Timeout = timeout

        address = ("localhost", 9956)

        self.Connection = Client(address, authkey=b"02812975")

        self.Readers = [self.Connection]


        '''
        object enable items added for better coding environment and simulation
        mode operation
        '''

        prog_only = 1
        if prog_only == 0:
            # ======== only for object programming
            # testing used temp instrument
            # need to become comment when the OBJ is finished
            import mcu_obj as mcu
            import inst_pkg_d as inst

            # add the libirary from Geroge
            import Scope_LE6100A as sco

            # initial the object and set to simulation mode
            pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, "off")
            pwr0.sim_inst = 0
            # initial the object and set to simulation mode
            met_v0 = inst.Met_34460(0.0001, 30, 0.000001, 2.5, 21)
            met_v0.sim_inst = 0
            loader_0 = inst.chroma_63600(1, 7, "CCL")
            loader_0.sim_inst = 0
            # mcu is also config as simulation mode
            mcu0 = mcu.MCU_control(0, 3)
            # using the main control book as default
            excel0 = par.excel_parameter("obj_main")
            src0 = inst.Keth_2440(0, 0, 24, "off", "CURR", 15)
            src0.sim_inst = 0
            met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
            met_i0.sim_inst = 0
            chamber0 = inst.chamber_su242(25, 10, "off", -45, 180, 0)
            chamber0.sim_inst = 0
            scope0 = sco.Scope_LE6100A("GPIB: 15", 0, 0)
            # ======== only for object programming

        # assign the input information to object variable
        self.excel_ini = excel0
        self.pwr_ini = pwr0
        self.loader_ini = loader_0
        self.met_v_ini = met_v0
        self.mcu_ini = mcu0
        self.src_ini = src0
        self.met_i_ini = met_i0
        self.chamber_ini = chamber0
        self.scope_ini = scope0
        # self.single_ini = single0


        # ====add few variable saving control setting
        # LPS505
        self.pwr_ch = 3

        # part_num for change sheet name
        self.part_num = 'NT50374'

        pass


    # ** put return value into "__return__"
    def call(self, func, *args, **kwargs):
        self.Connection.send((NAGSIGN_RPC, "__return__ = " + func, args, kwargs))

        for r in wait(self.Readers, timeout=self.Timeout):
            result = r.recv()

            if isinstance(result, dict) and "Error" in result:
                raise Exception(result["Error"])

            return result

    def run(self, codes, timeout=3):
        self.Connection.send((NAGSIGN_RPC, codes, (), {}))

        for r in wait(self.Readers, timeout=timeout):
            result = r.recv()

            if isinstance(result, dict) and "Error" in result:
                raise Exception(result["Error"])

            return result

        pass

    def eff_run(self, tag_name0='virtual'):
        '''
        input the related tage for efficiency operation
        need to setup tag in the before operation
        tag_name0 is string
        '''

        cmd_str_V5 = f"""
from GAutoVerify.PMU.Efficiency import Efficiency
Efficiency.showForm()
Efficiency.restoreByTag('{tag_name0}')
Efficiency.Run()
"""

        result = self.run(cmd_str_V5, timeout=1800)

        return result

    def buck_regulation_mix(self, mode0=0, setting_sel0=''):

        '''
        mode 0 => all, 1-LDO(L), 2-LDO_BK_on(LBO), 3-Buck_normal(BN), 4-Buck_usm(BU)
        setting_sel0 => choose related settings:
        need to have below setting files before operation
        'setting_sel' + '_L'

        this item pack all the HV buck eff and regulation testing together
        1. LDO only mode (EN2 tight high through MCU if there are EN2)
        2. LDO regulation scan when Buck turn on
        3. Buck normal mode (EN1 is normal voltage)
        4. Buck USM mode (EN1 is set to about 1.5V ect)

        231002: note~ file name cna change after find the way to adjust tag,
        if knowing how to adjust tage name and add part number,
        it's easier to operate with different part number

        '''
        setting_sel0 = str(setting_sel0)
        v_nor = 3.3
        v_usm = 1.5
        v_off = 0
        i_o_curr = 0.2
        mode_index = [ '_L', '_LBO', '_BN', '_BU' ]
        vio_index = [ v_off, v_nor, v_nor, v_usm ]

        # device initialization for this item
        self.pwr_ini.chg_out(v_off, i_o_curr, self.pwr_ch, 'off')


        # finished initialization
        time.sleep(1)

        if mode0 != 0 :
            # single operation

            # mode x_mode
            tag_set = setting_sel0 + mode_index[mode0]
            # finished EN1=L for LDO regulation

            self.pwr_ini.chg_out(
                vio_index[mode0],
                i_o_curr,
                self.pwr_ch,
                "on",
            )
            self.eff_run(tag_name0=tag_set)

            print(f'grace is still single?, the testing of {tag_set} is done')

            pass

        else:
            # all run

            c_mode = 4
            x_mode = 0

            while x_mode < c_mode :

                # mode x_mode
                tag_set = setting_sel0 + mode_index[x_mode]
                # finished EN1=L for LDO regulation

                self.pwr_ini.chg_out(
                    vio_index[x_mode],
                    i_o_curr,
                    self.pwr_ch,
                    "on",
                )

                self.eff_run(tag_name0=tag_set)

                # find the sheet and copy to the result temp
                sh_anme = f'<{tag_set}> #1'
                self.find_sh_all_books(sheet_name=sh_anme)

                print(f'all in one, the testing of {tag_set} is done')
                x_mode = x_mode + 1

            pass

        # get the second round of efficiency and regulstion

        # process the repo manager for merge the result together

        pass

    def find_sh_all_books(self, sheet_name=''):
        # 连接到 Excel 应用程序
        # app = xw.App(visible=False)  # 如果不需要可见 Excel，请设置 visible=False
        app2 = xw.apps

        # 获取所有已打开的工作簿
        workbooks = xw.books
        print(workbooks)
        print(app2)
        for app in app2:
            workbooks = app.books
            # 遍历所有已打开的工作簿
            for workbook in workbooks:
                # 遍历当前工作簿中的所有工作表
                for sheet in workbook.sheets:
                    if sheet.name == sheet_name:
                        # 找到匹配的工作表
                        print(f"在工作簿 '{workbook.name}' 中找到匹配的工作表：{sheet_name}")
                        # 在这里可以执行其他操作，如读取或修改工作表内容
                        '''
                        copy this sheet to the result book, it should be able to find the
                        result book and reference sheet for copying index during the program operation
                        '''
                        print(app.name)
                        # try to find the sheet, get the full trace,
                        # open books in same app
                        # and you can copy the sheet without pain
                        #


                        break
                    pass
                pass
            pass

        # 关闭 Excel 应用程序
        # app.quit()

        # return  new_sh




if __name__ == "__main__":
    import inst_pkg_d as inst

    # for the excel sheet control
    # only the main program use this method to separate the excel and program
    # since here will include the loop structure and overall control of report generation
    import sheet_ctrl_main_obj as sh
    import parameter_load_obj as para

    # define the excel object
    excel_m = para.excel_parameter(str(sh.file_setting))

    pwr_m = inst.LPS_505N(
    excel_m.pwr_vset,
    excel_m.pwr_iset,
    excel_m.pwr_act_ch,
    excel_m.pwr_supply_addr,
    excel_m.pwr_ini_state,)

    pwr_m.sim_inst = 0

    # original version of definition => without instrument object input
    # NAGui = NAGuiRPC()

    NAGui = NAGuiRPC(pwr0=pwr_m)

    # # 1-line function - method 1
    # result = NAGui.call('GI2C.read(0x9E, 0x00, 1)')
    # print('GI2C.read(0x9E, 0x00, 1) > ', result)

    # # 1-line function - method 2
    # result = NAGui.call('GI2C.read', 0x9E, 0x00, 1)
    # print('GI2C.read(0x9E, 0x00, 1) > ', result)

    test_index = 2.5

    # codes
    code = """

from GAutoVerify.PMU.Efficiency import Efficiency

Efficiency.showForm()

Efficiency.restoreByTag('virtual')

Efficiency.Run()

"""

    if test_index == 0:
        # default setting provide from Geroge

        result = NAGui.run(code, timeout=1800)

        '''
        the comments for using run function in RPC (will Grace know? who knows~ XD)
        first to check on the trace from main.py of GPL_V5
        here is the tarce of main.py: (also check from the git trace)
        D_GPL_tool_GPL_NAGui v5_NAGui py310_NAGui

        second is to import the obj from related folder (import the py file)
        "from GAutoVerify.PMU.Efficiency import Efficiency"
        GAutoVerify => the folder with main.py
        Efficiency.py is the related object, refer to the definition of object in file:
        Efficiency = Auto_PmuEfficiency() => "Efficiency" is the operation object

        "Efficiency.showForm()" => show the operation form of this item
        "Efficiency.restoreByTag('virtual')" => choolse the setting file
        "Efficiency.Run()" => start the verification

        note that the "restoreByTag" function is from previous object:
        "class Auto_PmuEfficiency(GTagForm, Ui_Form_CommonAV):"
        GTagForm => is where "restoreByTag" function from

        '''

        pass

    elif test_index == 1:
        # need to define new code for remote operation

        cmd_str_V5 = """
        from GAutoVerify.PMU.Efficiency import Efficiency
        Efficiency.showForm()
        Efficiency.restoreByTag('virtual')
        Efficiency.Run()
        """

        '''
        231006 new add comments
        double return can get variable from G_RPC, from GPL_V5
        __return__
        return format check with George
        '''

        result = NAGui.run(code, timeout=1800)

        pass

    elif test_index == 2:




        # define the tag file for V5
        tag_name ='virtual'
        tag_name2 = 'virtual1'

        # adjust MCU mode before efficiency operation
        # LDO only or AOD mode
        mcu_m.pmic_mode(3)
        NAGui.eff_run(tag_name0=tag_name)

        # normal mode
        mcu_m.pmic_mode(4)
        NAGui.eff_run(tag_name0=tag_name2)



        pass

    elif test_index == 2.5 :

        NAGui.part_num = 'SY8388C3'
        report_trace = ''
        target_sheet = ''


        NAGui.buck_regulation_mix(mode0=0, setting_sel0='virtual')

        pass

    elif test_index == 3 :
        '''
        call GPattern for pattern gen application
        refer to the pattern gen example:

        example:
        GPattern.run(0)             #run first pattern(index = 0)
        GPattern.run('Pattern1')    #run pattern or PGCB named 'Pattern1'

        '''


        pass

    elif test_index == 4 :
        # pyhton testing for xlwings




        pass
