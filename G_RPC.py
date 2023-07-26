from multiprocessing.connection import Client, wait

NAGSIGN_RPC = "##NAGRPC##"


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


if __name__ == "__main__":
    NAGui = NAGuiRPC()

    # # 1-line function - method 1
    # result = NAGui.call('GI2C.read(0x9E, 0x00, 1)')
    # print('GI2C.read(0x9E, 0x00, 1) > ', result)

    # # 1-line function - method 2
    # result = NAGui.call('GI2C.read', 0x9E, 0x00, 1)
    # print('GI2C.read(0x9E, 0x00, 1) > ', result)

    # codes
    code = """

from GAutoVerify.PMU.Efficiency import Efficiency

Efficiency.showForm()

Efficiency.restoreByTag('virtual')

Efficiency.Run()

"""

    result = NAGui.run(code, timeout=1800)
