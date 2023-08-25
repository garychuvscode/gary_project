from multiprocessing.connection import Client, wait

import JIGM3 as mcu_g

mcu_m = mcu_g.JIGM3(sim_mcu0=0)
print("JIGM3 MCU selected for Grace")

NAGSIGN_RPC = "##NAGRPC##"
# fmt: off

class NAGuiRPC:
    def __init__(self, timeout=3.0):
        self.Timeout = timeout

        address = ("localhost", 9956)

        self.Connection = Client(address, authkey=b"02812975")

        self.Readers = [self.Connection]

    # ** put return value into "__return__"
    def call(self, func, *args, **kwargs):
        self.Connection.send((NAGSIGN_RPC, "__return__ = " + func, args, kwargs))

        for r in wait(self.Readers, timeout=self.Timeout):
            result = r.recv()

            if isinstance(result, dict) and "Error" in result:
                raise Exception(result["Error"])

            return result

    def run(self, codes, timeout=3.0):
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


if __name__ == "__main__":
    NAGui = NAGuiRPC()

    # # 1-line function - method 1
    # result = NAGui.call('GI2C.read(0x9E, 0x00, 1)')
    # print('GI2C.read(0x9E, 0x00, 1) > ', result)

    # # 1-line function - method 2
    # result = NAGui.call('GI2C.read', 0x9E, 0x00, 1)
    # print('GI2C.read(0x9E, 0x00, 1) > ', result)

    test_index = 2

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

        pass

    elif test_index == 1:
        # need to define new code for remote operation

        cmd_str_V5 = """
        from GAutoVerify.PMU.Efficiency import Efficiency
        Efficiency.showForm()
        Efficiency.restoreByTag('virtual')
        Efficiency.Run()
        """

        result = NAGui.run(code, timeout=1800)

        pass

    elif test_index == 2:

        #  this item can be used to scan the LDO load regulation


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
