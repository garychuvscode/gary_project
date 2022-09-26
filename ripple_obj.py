# this file is used to generate the ripple testing file for
# verification object
# need to input the excel obj to reference the format gen sheet and related index
# different pages need to use different index loaded


class ripple_test ():

    def __init__(self, excel0):

        # # ======== only for object programming
        # # testing used temp instrument
        # # need to become comment when the OBJ is finished
        # import mcu_obj as mcu
        # import inst_pkg_d as inst
        # # initial the object and set to simulation mode
        # pwr0 = inst.LPS_505N(3.7, 0.5, 3, 1, 'off')
        # pwr0.sim_inst = 0
        # # initial the object and set to simulation mode
        # met_v0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 21)
        # met_v0.sim_inst = 0
        # loader_0 = inst.chroma_63600(1, 7, 'CCL')
        # loader_0.sim_inst = 0
        # # mcu is also config as simulation mode
        # mcu0 = mcu.MCU_control(0, 3)
        # # using the main control book as default
        # excel0 = par.excel_parameter('obj_main')
        # src0 = inst.Keth_2440(0, 0, 24, 'off', 'CURR', 15)
        # src0.sim_inst = 0
        # met_i0 = inst.Met_34460(0.0001, 7, 0.000001, 2.5, 20)
        # met_i0.sim_inst = 0
        # chamber0 = inst.chamber_su242(25, 10, 'off', -45, 180, 0)
        # chamber0.sim_inst = 0
        # # ======== only for object programming

        # assign the input information to object variable
        self.excel_ini = excel0

        pass
