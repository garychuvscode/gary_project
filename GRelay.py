
from .Form_GRelay_ui           import *
from .GRelayUnit               import GRelayUnit

from NAGlib.GTagForm                    import *
from NAGlib.GRegCtrl.FrameUnit          import *
from NAGlib.GSettingManager.GSavedDict  import GSavedDict

from NAGlib.GJIGManager.GJIGManager     import GJIGManager   

import re


class HLabel(QWidget):
    mouseLeftClicked = Signal(str)
    mouseRightClicked = Signal(str)

    def __init__(self, text=None):
        super(self.__class__, self).__init__()
        self.text = text

    def sizeHint(self):
        w,  h = FrameUnit.getHintSize(self.text)
        return QSize(w, h)

    def paintEvent(self, event):
        try:
            p = QPainter()
            p.begin(self)
            p.setPen( QColor(0xEEEEEE), )
            p.setBrush( QColor(0xEEEEEE) )
            p.setFont(FrameUnit.getFitFont(self.text, self.rect().width(), self.rect().height()))
            p.drawText(self.rect(), Qt.AlignVCenter | Qt.AlignLeft, self.text)  
        finally:
            p.end()
     
    def mousePressEvent(self, event) :     
        if event.button() == Qt.RightButton:
            self.mouseRightClicked.emit(self.text)
        else :
            self.mouseLeftClicked.emit(self.text)


class VLabel(QWidget):
    mouseLeftClicked = Signal(str)
    mouseRightClicked = Signal(str)

    def __init__(self, text=None):
        super(self.__class__, self).__init__()
        self.text = text

    def sizeHint(self):
        w,  h = FrameUnit.getHintSize(self.text)
        return QSize(h, w)

    def paintEvent(self, event):
        try:
            p = QPainter()
            p.begin(self)
            p.setPen( QColor(0xEEEEEE) )
            p.setBrush( QColor(0xEEEEEE) )
            p.setFont(FrameUnit.getFitFont(self.text, self.rect().height()-2, self.rect().width()))
            p.rotate(90)
            p.translate(0, -1 * self.rect().width())
            rect = QRect(self.rect().x(), self.rect().y(), self.rect().height(), self.rect().width())
            p.drawText(rect,  Qt.AlignVCenter | Qt.AlignLeft, self.text)  
        finally:
            p.end()

    def mousePressEvent(self, event) :     
        if event.button() == Qt.RightButton:
            self.mouseRightClicked.emit(self.text)
        else :
            self.mouseLeftClicked.emit(self.text)


class NamingDialog(QDialog):
    LastBusName = ''
    lastPinName = ''

    def __init__(self, relay_code, bus_name, pin_name, parent=None):
        super().__init__(parent)

        self.setWindowFlags(Qt.CustomizeWindowHint | Qt.Dialog )

        self.layout = QGridLayout()

        self.layout.addWidget(QLabel(relay_code),       0, 0)
        self.layout.addWidget(QLabel(f'new name'),      0, 1)
        self.layout.addWidget(QLabel(f'(original)'),    0, 2, alignment = Qt.AlignLeft)

        self.layout.addWidget(QLabel(f'PinName:'),      1, 0)
        self.Edit_PinName = QLineEdit(NamingDialog.lastPinName)
        self.layout.addWidget(self.Edit_PinName,        1, 1)
        self.layout.addWidget(QLabel(f'({pin_name})'),  1, 2, alignment = Qt.AlignLeft)
        self.Edit_PinName.selectAll()        

        self.layout.addWidget(QLabel(f'BusName:'),      2, 0)
        self.Edit_BusName = QLineEdit(NamingDialog.LastBusName)
        self.layout.addWidget(self.Edit_BusName,        2, 1)
        self.layout.addWidget(QLabel(f'({bus_name})'),  2, 2, alignment = Qt.AlignLeft)


        self.buttonok = QDialogButtonBox(QDialogButtonBox.Ok)
        self.buttonok.clicked.connect(self.accept)
        self.layout.addWidget(self.buttonok,            3, 1)  
        self.buttonCancel = QDialogButtonBox(QDialogButtonBox.Cancel)
        self.buttonCancel.clicked.connect(self.reject)
        self.layout.addWidget(self.buttonCancel,        3, 2, alignment = Qt.AlignLeft) 

        self.setLayout(self.layout)

    def accept(self):
        super().accept()

        NamingDialog.LastBusName = self.Edit_BusName.text()
        NamingDialog.lastPinName = self.Edit_PinName.text()


class NamingAllDialog(QDialog):
    LastName = ''

    def __init__(self, current_name, parent=None):
        super().__init__(parent)

        self.setWindowFlags(Qt.CustomizeWindowHint | Qt.Dialog )

        self.layout = QGridLayout()

        self.layout.addWidget(QLabel(f'old name:'),    0, 0, alignment = Qt.AlignCenter)
        self.layout.addWidget(QLabel(f'new name:'),    0, 1, alignment = Qt.AlignCenter)

        self.layout.addWidget(QLabel(f'({current_name})'),  1, 0, alignment = Qt.AlignCenter)
        self.Edit_NewName = QLineEdit(NamingAllDialog.LastName)
        self.layout.addWidget(self.Edit_NewName,            1, 1, alignment = Qt.AlignCenter)
        
        self.Edit_NewName.selectAll()        


        self.buttonok = QDialogButtonBox(QDialogButtonBox.Ok)
        self.buttonok.clicked.connect(self.accept)
        self.layout.addWidget(self.buttonok,            2, 0)  
        self.buttonCancel = QDialogButtonBox(QDialogButtonBox.Cancel)
        self.buttonCancel.clicked.connect(self.reject)
        self.layout.addWidget(self.buttonCancel,        2, 1, alignment = Qt.AlignLeft) 

        self.setLayout(self.layout)


    def accept(self):
        super().accept()

        NamingAllDialog.LastName = self.Edit_NewName.text()





# -----------------------------------------------------------------------------------------------------------------
# --  GRelay ------------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------
class Relay(GTagForm, Ui_Form_GRelay):

    relayResetSignal = Signal()
    relayStatusUpdated = Signal(list, bool)

    @run_in_main()
    def __init__(self, parent = None) :
        super().__init__(parent)
        self.setupUi(self)

        self.setIconPath(u":/image/resources/ICON_Relay.svg")

        self.OnlyOneTagForm = True

        for i in range(255) :
            setattr(self, f'P{i+1}',   (i + 1)*0x00000001)
        for i in range(15) :
            setattr(self, f'BUS{i+1}', (i + 1)*0x00000100)

        self.Btn_ToggleSlider.clicked.connect(self.togglePanel)
        self.Btn_DisattachAll.clicked.connect(self.disattachAll)
        self.relayStatusUpdated.connect(self.updateRelayStatus)
        self.relayResetSignal.connect(self.reset)
        self.Edit_BusNum.returnPressed.connect(self.onBusPinNumChanged)
        self.Edit_PinNum.returnPressed.connect(self.onBusPinNumChanged)
        self.Btn_DisattachAllResistors.clicked.connect(threadin_user_button(self.setNoLoad))
        self.Combo_TotalR.activated.connect(self.onTotalRSelected)

        self.SavedVars = GSavedDict({'BusNum' : 4, 'PinNum': 8})

        self.WgDoRelayConnected = None

        self.Layout_Tag.addWidget(self.Tag) 

        GSettingManager.restore(self)


    def restoreEnd(self): #Callback after GSettingManager.restore
        self.reset()
        self.resetRLoader()

        if self.WgDoRelayConnected is not None :
            self.WgDoRelayConnected['widget'].Label_Tag.setText(self.Tag.text())   

    
    def clearLayout(self, layout):
        for i in reversed(range(layout.count())): 
            layout.itemAt(i).widget().setParent(None)

    def reset(self):
        Neighbors = {}
        self.RelayCodeMap = {}
        self.SugarNameMap = {}

        bus_num = self.SavedVars['BusNum']
        pin_num = self.SavedVars['PinNum']
        self.Edit_BusNum.setText(str(bus_num))
        self.Edit_PinNum.setText(str(pin_num))

        layout = self.Layout_Relays
        self.clearLayout(layout)

        for bus_i in range(1, bus_num+1) :        
            for pin_i in range(1, pin_num+1) :

                bus = f'BUS{bus_i}'                
                pin = f'P{pin_i}'
                relay_code = f'0x0000{bus_i:02X}{pin_i:02X}'

                rvars = self.SavedVars.get(relay_code, {'BusName':bus, 'PinName':pin})  

                g = GRelayUnit(bus_i, pin_i, rvars['BusName'], rvars['PinName'])
                g.mouseClicked.connect(self.onRelayClicked)
                g.mouseRightClicked.connect(self.askUserType)

                col = 2 *(pin_num + 1 - pin_i)
                row = 2 *(bus_num + 1 - bus_i) 
                layout.addWidget(g, row, col)

                if (rvars['BusName']  != Neighbors.get(bus)) :
                    label = HLabel(rvars['BusName'] )
                    label.mouseRightClicked.connect(self.askUserTypeAll)
                    layout.addWidget(label, row,     col + 1)

                if (rvars['PinName']  != Neighbors.get(pin)) :
                    label = VLabel(rvars['PinName'] )
                    label.mouseRightClicked.connect(self.askUserTypeAll)
                    layout.addWidget(label, row + 1, col)

                Neighbors[bus] = rvars['BusName']
                Neighbors[pin] = rvars['PinName']

                self.RelayCodeMap[relay_code] = g
                self.SugarNameMap[relay_code] = relay_code
                self.SugarNameMap[f"{rvars['BusName']}+{rvars['PinName']}"] = relay_code
                self.SugarNameMap[f"{rvars['PinName']}+{rvars['BusName']}"] = relay_code


    def updateRelayStatus(self, relay_codes, connected) :
        for relay_code in relay_codes :
            relay_unit = self.RelayCodeMap[relay_code]
            relay_unit.setConnected(connected)

    # -----------------------------------------------------------------------------------------------------------------
    # --  Slider Panel -------------------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------
    def togglePanel(self):
        extend_height = self.Frame_SliderPanel.sizeHint().height()
        #print(f"{self.Frame_SliderPanel.sizeHint()=}")
        if self.Frame_SliderPanel.maximumHeight() < 16 :
            self.anilfm = QPropertyAnimation(self.Frame_SliderPanel, b"maximumHeight")
            self.anilfm.setEasingCurve(QEasingCurve.OutExpo)
            self.anilfm.setDuration(500)
            self.anilfm.setStartValue(2)
            self.anilfm.setEndValue(extend_height)
            self.anilfm.start() 

            self.anihei = QPropertyAnimation(self, b"geometry")
            self.anihei.setEasingCurve(QEasingCurve.OutExpo)
            self.anihei.setDuration(500)
            self.anihei.setStartValue(self.geometry())
            self.anihei.setEndValue(self.geometry().adjusted(0, 0, 0, extend_height))

            self.anigroup = QParallelAnimationGroup()
            self.anigroup.addAnimation(self.anilfm)
            self.anigroup.addAnimation(self.anihei)
            self.anigroup.start()  

            icon = QIcon()
            icon.addFile(u":/IconTool/resources/ICON_ArrowUp.svg", QSize(), QIcon.Normal, QIcon.Off)
            self.Btn_ToggleSlider.setIcon(icon)
        else :
            self.anilfm = QPropertyAnimation(self.Frame_SliderPanel, b"maximumHeight")
            self.anilfm.setEasingCurve(QEasingCurve.OutExpo)
            self.anilfm.setDuration(500)
            self.anilfm.setStartValue(extend_height)
            self.anilfm.setEndValue(2)
            self.anilfm.start() 

            self.anihei = QPropertyAnimation(self, b"geometry")
            self.anihei.setEasingCurve(QEasingCurve.OutExpo)
            self.anihei.setDuration(500)
            self.anihei.setStartValue(self.geometry())
            self.anihei.setEndValue(self.geometry().adjusted(0, 0, 0, -extend_height))            

            self.anigroup = QParallelAnimationGroup()
            self.anigroup.addAnimation(self.anilfm)
            self.anigroup.addAnimation(self.anihei)
            self.anigroup.start()  

            icon = QIcon()
            icon.addFile(u":/IconTool/resources/ICON_ArrowDown.svg", QSize(), QIcon.Normal, QIcon.Off)
            self.Btn_ToggleSlider.setIcon(icon)


    # -----------------------------------------------------------------------------------------------------------------
    # --   -------------------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------

    @threadin_user_button
    @log_error
    def onRelayClicked(self, relay_code):
        relay_unit = self.RelayCodeMap[relay_code]
        if relay_unit.connected() :
            self.disattach(relay_code)
        else :
            self.attach(relay_code)

    def askUserType(self, relay_code):
        relay = self.RelayCodeMap[relay_code]
        dlg = NamingDialog(relay.RelayCode, relay.BusName, relay.PinName, self)
        ok = dlg.exec_()
        if ok :            
            self.SavedVars[relay_code] = {'BusName':dlg.Edit_BusName.text(), 'PinName':dlg.Edit_PinName.text()}
            self.reset()


    def askUserTypeAll(self, text):
        dlg = NamingAllDialog(text, self)
        ok = dlg.exec_()
        if ok :   
            newname = dlg.Edit_NewName.text()
            for relay_code, relay in self.RelayCodeMap.items() :
                    if relay.BusName == text :
                        self.SavedVars[relay_code] = {'BusName':newname,       'PinName':relay.PinName}
                    if relay.PinName == text :
                        self.SavedVars[relay_code] = {'BusName':relay.BusName, 'PinName':newname}

            self.reset()



    @log_error
    def onBusPinNumChanged(self):
        self.SavedVars['BusNum'] = int(self.Edit_BusNum.text())
        self.SavedVars['PinNum'] = int(self.Edit_PinNum.text())
        self.relayResetSignal.emit()

    @threadin_user_button
    def disattachAll(self):
        self.reattach()
    # -----------------------------------------------------------------------------------------------------------------
    # --  RLoad -------------------------------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------
    def resetRLoader(self):
        self.Combo_TotalR.clear()
        self.genResistTB() 
        l = sorted(self.ResistTB.keys(), reverse = True)
        self.Combo_TotalR.addItems( [str(i) for i in l])

    def onTotalRSelected(self):
        self.setNoLoad()

        resist = int(self.Combo_TotalR.currentText())
        choose_resist, relays = self.getFitResist(resist , self.ResistTB)  

        self.attach(relays)

        self.setStatus( f'[RLoader] NeedR={resist} closestR={choose_resist}',timeout_ms = 5000)     


    def testRload(self):
        self.ResistTB = self.genResistTB()   
        print(self.ResistTB)  


    def genResistTB(self):
        self.ResistTB = {}
        resistor = []
        self.ResistRcodes = []
        for relay_code, relay in self.RelayCodeMap.items() :
            match = re.search(r"R(\d+)", relay.PinName)
            if match is not None:
                resistor.append(int(match.group(1)))
                self.ResistRcodes.append(relay_code)

        p = [ (0, i+1) for i in range(len(resistor)) ]

        from itertools import product
        for indices in product(*p):

            conductance = 0
            RiPicks = []
            for RiPicked in indices:
                if RiPicked == 0 : continue

                conductance += 1/resistor[RiPicked-1]
                RiPicks.append(self.ResistRcodes[RiPicked-1])

            if conductance > 0 :
                resist = int(1/conductance)
                self.ResistTB[resist] = RiPicks


    def getFitResist(self, resist, resist_table):
        resist_keys = list(resist_table.keys())
        closekey = min(resist_keys, key= lambda x : (x - resist) if (x - resist) >= 0 else 1e9)

        return closekey, resist_table[closekey]

    # -----------------------------------------------------------------------------------------------------------------
    def connectWidget(self, widget):
        self.WgDoRelayConnected = {'widget': widget}

    def jig(self):
        return GJIGManager.MyJIG('GRelay', prefer=['MxN16R', 'ControlBoardCR'], shared=self.Radio_ShareJIG.isChecked())

    # -----------------------------------------------------------------------------------------------------------------
    # --  EXPORT (thread-save)-----------------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------
    def toRelayCodes(self, *args):
        """
        RelayCodes
        ================================================================
        [RelayCodes Example] 
        :param *relays: ex. 
        :                   GRelay.reattach(0x00000101, 0x00000102, 0x00000201)
        :                   GRelay.reattach([0x00000101, 0x00000102, 0x00000201])
        :                   GRelay.reattach(0x00000101, [0x00000102, 0x00000201])
        :                   GRelay.reattach("BUS1+P1", "AVDD+METER1")
        """      
        relays_codes = []
        l = []        
        for arg in args:
            if isinstance(arg, list) or isinstance(arg, tuple):
                l.extend(arg)
            else :
                l.append(arg)

        for o in l :
            if isinstance(o, str):
                relaycode = self.SugarNameMap.get(o.replace(' ', ''))
                if relaycode is not None :
                    relays_codes.append(relaycode)
                else :
                    for sugar_name, relaycode in self.SugarNameMap.items():
                        if re.search(o, sugar_name) is not None :
                            relays_codes.append(relaycode)
                    
                    if len(relays_codes) == 0 :
                        raise Exception(f'<>< GRelay ><> ERROR : cannot find Relay [{o}]')

            if isinstance(o, int):
                relays_codes.append(f'0x{o:08X}')

        return relays_codes
    
      
    def reattach(self, *relays):
        """
        GRelay.reattach(*relays) -> None
        ================================================================
        [GRelay disconnect All then connect the relays] 
        :param *relays: see RelayCode
        :return: None
        """
        logging.info(f'relays = {relays}')
        relays_codes = self.toRelayCodes(*relays)

        result = self.jig().ezCommand(f"return  mcu.autotrim.reconnect({', '.join(relays_codes)})") 

        if result[0] > 0  :
            self.setStatus(f"ERROR MCU Code:{result[1]}", GBaseForm.StatusErrorStyle)
        else :
            self.setStatus(f"reattach[{','.join(relays_codes)}]")
            self.relayStatusUpdated.emit(list(self.RelayCodeMap.keys()), False)
            self.relayStatusUpdated.emit(relays_codes, True)

    def attach(self, *relays):
        """
        GRelay.attach(*relays) -> None
        ================================================================
        [GRelay connect the relays] 
        :param *relays: see RelayCode
        :return: None
        """
        logging.info(f'relays = {relays}')
        relays_codes = self.toRelayCodes(*relays)

        result = self.jig().ezCommand(f"return  mcu.autotrim.connect({', '.join(relays_codes)})") 

        if result[0] > 0  :
            self.setStatus(f"ERROR MCU Code:{result[1]}", GBaseForm.StatusErrorStyle)
        else :
            self.setStatus(f"attach[{','.join(relays_codes)}]")
            self.relayStatusUpdated.emit(relays_codes, True)


    def disattach(self, *relays):
        """
        GRelay.reconnect(*relays) -> error_id
        ================================================================
        [GRelay UI read i2c] 
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param len: length for read.
        :return: error_id
        """
        logging.info(f'relays = {relays}')
        relays_codes = self.toRelayCodes(*relays)

        result = self.jig().ezCommand(f"return  mcu.autotrim.disconnect({', '.join(relays_codes)})") 

        if result[0] > 0  :
            self.setStatus(f"ERROR MCU Code:{result[1]}", GBaseForm.StatusErrorStyle)
        else :
            self.setStatus(f"disattach[{','.join(relays_codes)}]")
            self.relayStatusUpdated.emit(relays_codes, False)

    # advanced--------------------------
    def setLoad(self, resist):
        choose_resist, relays = self.getFitResist(resist , self.ResistTB)  

        self.setNoLoad()
        self.attach(relays)

        self.setStatus( f'[RLoader] NeedR={resist} closestR={choose_resist}',timeout_ms = 5000) 

        return choose_resist 

    def setNoLoad(self):
        self.disattach(self.ResistRcodes)




GRelay = Relay('GRelay') 