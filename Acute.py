# -*- coding: utf-8 -*-
from ctypes import *
import ctypes
import logging,os, re
import time
from pathlib import Path

from .GInst import *


#AqLAVISA constants
#Error code
AQVI_DLL_NOT_READY = 9999

#VISA Get Command Result
PROC_OK_WOUT_RETURN = 0
PROC_OK_WITH_RETURN = 1
PROC_FAIL = 2
PROC_PENDING = 3

#Environment Settings
AqLAVISA_DLL_Directory = str(Path().resolve() / 'dll' / 'AqLAVISA64.dll')

AcuteLinkType = {
    'TL.exe (TL3000+)'      : 0,
    'BF_LA3K.exe (LA3000+)' : 1,
    'TBA.exe'               : 2,
    'MSO.exe'               : 3,
}

class Acute(GInst):    

    def __init__(self, link, ch) :
        super().__init__()    

        self.DllDirectory = AqLAVISA_DLL_Directory
        self.LaDll = ctypes.WinDLL(AqLAVISA_DLL_Directory)
        self.LaDll.viWrite.argtypes = [ctypes.c_char_p]
        self.LaDll.viRead.argtypes = [ctypes.c_char_p, ctypes.c_int]           

        apptype = AcuteLinkType[link]

        if (self.viSelectAppType(apptype) == 0):
            raise Exception("viSelectAppType")
            return

        if (self.viOpenRM() == 0):
            raise Exception("viOpenRM")
      

    def __del__(self):
        self.viCloseRM()


    #Select Target Application Type
    def viSelectAppType(self, AppType):
        RetryCount = 10 
        while (self.LaDll.viSelectAppType(AppType) == 0):
            if (self.LaDll.viErrCode() != AQVI_DLL_NOT_READY): #Error code = 9999, DLL still initializing
                logging.warning ("viSelectAppType Error... ", self.LaDll.viErrCode())
                return False
            if (RetryCount == 0):
                logging.warning ("viSelectAppType Timeout... ", self.LaDll.viErrCode())
                return False
            RetryCount -= 1
            time.sleep(0.1) #Wait 100ms for each loop	
        return True
        
                
    #Establish S/W connection
    def viOpenRM(self):
        RetryCount = 10 
        while (self.LaDll.viOpenRM() == 0):
            if (self.LaDll.viErrCode() != AQVI_DLL_NOT_READY): #Error code = 9999, DLL still initializing
                logging.warning ("viOpenRM Error... ", self.LaDll.viErrCode())
                return False
            if (RetryCount == 0):
                logging.warning ("viOpenRM Timeout... ", self.LaDll.viErrCode())
                return False
            RetryCount -= 1
            time.sleep(0.1) #Wait 100ms for each loop	
        return True

    #Read Error Code
    def viErrCode(self):
        return self.LaDll.viErrCode()

    #Disconnect from S/W
    def viCloseRM(self):
        if (self.LaDll.viCloseRM() == 0):
            logging.warning ("viCloseRM Error... ", self.LaDll.viErrCode())


    #Send Command to S/W
    def viWrite(self, cmd, timeout = 10):     
        """
        logic.viWrite(cmd, timeout = 10) -> None
        ================================================================
        [Acute LA viWrite] 
        :param cmd: command
        :timeout = 
        :return: None
        """            
        cmd = cmd.encode('utf-8')
        if (self.LaDll.viWrite(ctypes.c_char_p(cmd)) == 0):
            return False
        TimerInterval = 0.1 #100ms
        RetryCount = timeout / TimerInterval
        CmdResult = self.LaDll.viGetCommandResult()
        while (CmdResult == PROC_PENDING and RetryCount > 0):
            time.sleep(TimerInterval) #Wait 100ms for command processing
            CmdResult = self.LaDll.viGetCommandResult()
            RetryCount -= 1
            
        if (CmdResult == PROC_FAIL):
            logging.warning("viWrite Error...", self.LaDll.viErrCode(), cmd)
            return False
        if (CmdResult == PROC_PENDING):
            logging.warning("viWrite Timeout...", self.LaDll.viErrCode(), cmd)
            return False
        return True

    #Read Command Return from S/W
    def viRead(self):
        BufSize = 1024
        ReadBuf = ctypes.create_string_buffer(BufSize)
        if (self.LaDll.viRead(ReadBuf, BufSize) == 0):
            logging.warning ("viRead Error... ", self.LaDll.viErrCode())
            return ""
        else: 
            return ReadBuf.value.decode("utf-8")

    def viQuery(self, command, timeout = 10):
        """
        logic.viQuery(command) -> ReadBuf.value
        ================================================================
        [Acute LA viQuery] 
        :param cmd: command
        :timeout =         
        :return: ReadBuf.value
        """           

        self.viWrite(command)
        return self.viRead()


    def printScreenToPC(self, path = None):
        """
        logic.printScreenToPC(path) -> None
        ================================================================
        [print Acute LA Screen and save to path] 
        :param path: file path or None for default_path 
        :return: None
        """   
        from pathlib import Path
        from NAGlib.Gsys.Gsys       import Gsys

        while self.viQuery("*LA:CAPTURE:STATUS?") != "Analysis Finished" :
            Gsys.msleep(100)        

        if path is None :
            from datetime import datetime     
            try :
                wavepath = f'D:\\@AutoVerify\\{datetime.now().strftime("%Y%m%d")}\\Waveforms'
                Path(wavepath).mkdir(parents=True, exist_ok=True)
            except :
                wavepath = wavepath.replace("D:\\@AutoVerify", "C:\\@AutoVerify")
                Path(wavepath).mkdir(parents=True, exist_ok=True)     

            pure_path = f"{wavepath}\\{datetime.now().strftime('%Y%m%d_%H%M%S%f')[:-5]}.png"   
        else :
            pure_path = os.path.splitext(path)[0] + ".png"

        self.viWrite(f"*FILE:SAVEIMAGE {pure_path}")
        return Path(pure_path)


    def triggerSingle(self, sync = True) :
        """
        logic.triggerSingle(sync = True) -> None
        ================================================================
        [logic trigger Single] 
        :param sync(option): True > wait for Pre-Trigger done
        :return: None

        example:
            logic.triggerSingle()      
        """      
        from NAGlib.Gsys.Gsys       import Gsys
        try :     
            self.viWrite("*LA:CAPTURE:START")

            while sync and self.viQuery("*LA:CAPTURE:STATUS?") in ['Capturing', 'Pre-Trigger'] :
                Gsys.msleep(100)

        except Exception as e :
            print('triggerSingle > ', type(e), str(e))





    def waitTriggered(self, timeout_ms = 5000) :
        """
        logic.waitTriggered(timeout = 5) -> bool
        ================================================================
        [wait until scope is triggered] 
        :param timeout(option): timeout second
        :return: None
        """   

        import time
        interval_ms = 330

        for t in range(0, timeout_ms, interval_ms) :
            if self.viQuery("*LA:CAPTURE:TRIGGERED?") == '1':
                return True
            time.sleep(interval_ms/1000)

        return False


        



    # def triggerStop(self) :
    #     """
    #     scope.triggerSingle() -> None
    #     ================================================================
    #     [scope trigger Single] 
    #     :return: None
    #     """           
    #     self.writeVBS('app.Acquisition.TriggerMode = "Single"')
    #     self.writeVBS('app.ClearSweeps')  



# test 


def CaptureAndSaveReport():    

    LaVisaDll = Acute(1, '')    

    if (LaVisaDll.viSelectAppType(0) == 0):
        return

    print ("Initial software connection...")
    if (LaVisaDll.viOpenRM() == 0):
        return;    
        
    print ("Begin software capture...")
    if (LaVisaDll.viWrite(b"*LA:CAPTURE:START") == 0): #Begin Capture
        return
    
    print ("Wait for waveform capture...")
    time.sleep(10) #Wait 10s for waveform capture    

    print ("Stop software capture, and wait until software ready")
    if (LaVisaDll.viWrite(b"*LA:CAPTURE:STOP") == 0): #Stop Capture
        return
        
    while (1): #wait until capture finished
        if (LaVisaDll.viWrite(b"*STB?") == 0): #Status check            
            return 
        ReadString = LaVisaDll.viRead();
        if (ReadString == ""):
            return
        if (ReadString == b'1'):
            break
        
    print ("Begin save captured data as report...")
    SaveCSVWaitTime = 30 #wait 30s for file saving
    # if (LaVisaDll.viWrite(b"*PA:REPORT:SAVE " + ReportCsvSaveDirectory, SaveCSVWaitTime) == 0): #Save Report
    #     return
    # while (1): #wait until report save finished
    #     if (LaVisaDll.viWrite(b"*STB?") == 0): #Status check
    #         return
    #     ReadString = LaVisaDll.viRead();
    #     if (ReadString == ""):            
    #         return
    #     if (ReadString == b'1'):
    #         break

    print ("Disconnect from software...")
    LaVisaDll.viCloseRM() 
    print ("All Finished!")
    
def main():
    CaptureAndSaveReport()
   
