
import win32com.client #import the pywin32 library
import pythoncom 
import logging,os

from .GInst import *


class Scope_LE6100A(GInst):
    ActiveDSO = None

    def __init__(self, link, ch) :
        super().__init__()        

        self.link = link
        self.ch = ch
        logging.debug(f'Initialize LecroyActiveDSO link={link}, ch={ch}')

        if Scope_LE6100A.ActiveDSO is None :
            pythoncom.CoInitialize()
            Scope_LE6100A.ActiveDSO = win32com.client.Dispatch("LeCroy.ActiveDSOCtrl.1") 
            pythoncom.CoInitialize()

        self.inst = Scope_LE6100A.ActiveDSO
        r = self.inst.MakeConnection(link)
        logging.debug(f"self.inst.MakeConnection(link) = {r}")

        if r == 0 :
            raise Exception('<>< Scope_LE6100A ><> LecroyActiveDSO link fail.')

        self.inst.SetRemoteLocal(1)
        self.inst.SetTimeout(20)



    def writeVBS(self, cmd):
        """
        scope.writeVBS(cmd) -> None
        ================================================================
        [write Lecroy VBS-Command to Scope] 
        :param cmd: lecory VBS command
        :return: None.
        """        
        logging.debug(f': {cmd}')
        self.inst.WriteString(f"VBS '{cmd}'", 1)

        if 0 == self.inst.WaitForOPC():
            raise Exception('LecroyActiveDSO writeVBS timeout or fail.')


    def readVBS(self, cmd):
        """
        scope.readVBS(cmd) -> string
        ================================================================
        [read string after write Lecroy VBS-Command to Scope] 
        :param cmd: lecory VBS command
        :return: scope return string.
        """     
        logging.debug(f': {cmd}')
        self.inst.WriteString(f"VBS? '{cmd}'", 1)

        return self.inst.ReadString(256) #reads a maximum of 256 bytes

    def readVBS_float(self, cmd):
        """
        scope.readVBS_float(cmd) -> float
        ================================================================
        [ get float value after write Lecroy VBS-Command to Scope] 
        :param cmd: lecory VBS command
        :return: float value.
        """    
        try :
            return float(self.readVBS(cmd))
        except:
            return -999

    def printScreenToPC(self, path):
        """
        scope.printScreenToPC(path) -> None
        ================================================================
        [print Lecroy Screen and save to path] 
        :param path: file path
        :return: None
        """   
        pure_path = os.path.splitext(path)[0]
        self.inst.StoreHardcopyToFile("PNG", "BCKG,WHITE,AREA,GRIDAREAONLY", pure_path + ".png")    

    def triggerSingle(self) :
        #logging.debug(f': {cmd}')
        #self.inst.WriteString(f"VBS '{cmd}'", 1)
        #self.writeVBS('app.Acquisition.TriggerMode = "Stopped"')
        self.writeVBS('app.Acquisition.TriggerMode = "Single"')
        self.writeVBS('app.ClearSweeps')  

        #self.inst.WriteString(f"VBS 'TriggerDetected = app.Acquisition.Acquire({timeout}, false)'", 1)
        #if 0 == self.inst.WaitForOPC():
        #    raise Exception('LecroyActiveDSO writeVBS timeout or fail.')

    def waitTriggered(self, timeout_ms = 5000) :
        """
        scope.waitTriggered(timeout = 5) -> bool
        ================================================================
        [wait until scope is triggered] 
        :param timeout(option): timeout second
        :return: None
        """   
        
        import time
        interval_ms = 330

        for t in range(0, timeout_ms, interval_ms) :
            if self.readVBS(f'return = app.Acquisition.TriggerMode') == 'Stopped':
                return True
            time.sleep(interval_ms/1000)

        return False
        # if 0 == self.inst.WaitForOPC():
        #     raise Exception('LecroyActiveDSO acquire fail.')

        # r = int(self.readVBS(f'return = TriggerDetected'))

        # if r :
        #     return True
        # else :
        #     raise Exception("LE6100A waitTriggered timeout!")


    def getFitScale(self, value, in_grid = 8) :
        """
        scope.getFitScale(value, in_grid) -> scale(float)
        ================================================================
        [get fit scale to make value view in in_grid] 
        :param value: 
        :param in_grid: 
        :return: None
        """  
        value_one_grid = value / in_grid

        # E = M * 10^N
        vestr = f'{value_one_grid:e}'
        ep = vestr.rfind("e")

        M = float(vestr[:ep])
        N = float(vestr[ep+1:])

        print(f"vestr={vestr}, M={M}, N={N}")
        if   M <= 1 : 
            vM = 1
        elif M <= 2 : 
            vM = 2
        elif M <= 5 : 
            vM = 5
        else :              
            vM = 10

        return vM * (10 ** N)


