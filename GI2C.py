# -*- coding: utf-8 -*-
from NAGlib.GBaseForm       import *
from .Form_I2C_ui           import *

from NAGlib.GJIGManager.GJIGManager  import GJIGManager    

from NAGlib.Except import I2CError

from PySide6.QtGui import QRegion

import logging, re

class I2CDataModel(QAbstractTableModel):
    clearlater = Signal()

    def __init__(self, lightcolor, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.LightColor = QColor(lightcolor)

        self.Datas  = [ [None]*16 for _ in range(16) ]
        self.Lights = [ [None]*16 for _ in range(16) ]


        self.clearlater.connect(self.clearlater_slot)

        self.cleartimer = QTimer(self) 
        self.cleartimer.setSingleShot(True)
        self.cleartimer.timeout.connect(self.clearLights)         

    def clearlater_slot(self):
        self.cleartimer.start(1500)     

    def clearLights(self):
        self.Lights = [ [None]*16 for _ in range(16) ]        
        self.dataChanged.emit(self.index(0, 0), self.index(16, 16), [Qt.BackgroundRole, Qt.ForegroundRole])   

    #-----------------------------------------------------------
    #-----------------------------------------------------------
    #-----------------------------------------------------------
    def rowCount(self, index):
        return 16

    def columnCount(self, index):
        return 16

    def headerData (self, section, direction, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:  
            if direction == Qt.Horizontal:
                return f"-{section:X}"

            if direction == Qt.Vertical:
                return f'{section:X}-'


    def data(self, index, role):
        if index.isValid():
            if role == Qt.DisplayRole:                
                return self.Datas[index.row()][index.column()]

            if role == Qt.BackgroundRole :
                return self.Lights[index.row()][index.column()]

            if role == Qt.ForegroundRole :
                return QColor('Black') if self.Lights[index.row()][index.column()] else None

            if role == Qt.TextAlignmentRole:
                return int(Qt.AlignCenter)


    #-----------------------------------------------------------
    #-----------------------------------------------------------
    #-----------------------------------------------------------
    def updateData(self, address, datas):
        min_row, min_col, max_row, max_col = 999, 999, -999, -999

        for i, data in enumerate(datas)  :
            row = int((address + i)/16)
            col = int((address + i)%16)

            self.Datas[row][col]  = f'{data:02X}'   if data is not None else None
            self.Lights[row][col] = self.LightColor if data is not None else None

            min_row = row if row < min_row else min_row
            max_row = row if row > max_row else max_row
            min_col = col if col < min_col else min_col
            max_col = col if col > max_col else max_col        
        
        self.dataChanged.emit(self.index(min_row, min_col), self.index(max_row, max_col), [Qt.DisplayRole, Qt.BackgroundRole, Qt.ForegroundRole])
        self.clearlater.emit()



class I2C(GBaseForm, Ui_Form_I2C):

    dataRxUpdated = Signal(int, list, str)
    dataTxUpdated = Signal(int, list, str)

    ErrorCode = { 'R/W Fail':1,  }

    I2CError = I2CError

    @run_in_main()
    def __init__(self, parent = None) :
        super().__init__(parent, progressbar=False)
        self.setupUi(self)

        self.setWindowIcon(QIcon(u":/image/resources/ICON_I2C.svg"))

        
        #self.initTable(self.Table_Rx)
        #self.initTable(self.Table_Tx)
        self.RxModel = I2CDataModel("lightgreen")
        self.Table_Rx.setModel(self.RxModel)
        self.Table_Rx.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.Table_Rx.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.TxModel = I2CDataModel("lightblue")
        self.Table_Tx.setModel(self.TxModel)
        self.Table_Tx.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.Table_Tx.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.Btn_Read.clicked.connect(self.readClicked)
        self.Btn_Write.clicked.connect(self.writeClicked)
        self.Btn_ClearCells.clicked.connect(self.clearCells)

        #self.dataUpdated.connect(self.highlight_cell)
        #self.dataRxUpdated.connect(lambda *args : self.highlight_cell(self.Table_Rx, *args))
        #self.dataTxUpdated.connect(lambda *args : self.highlight_cell(self.Table_Tx, *args))

        #GJIGManager.showOnlyOneForm(noshow = True)
        GSettingManager.restore(self)

        #self.clearColorCD = -1

        #self.timer = QTimer(self) 
        #self.timer.timeout.connect(self.timer_tick) 
        #self.timer.start(300) 





#     def initTable(self, table) :
#         table.setHorizontalHeaderLabels( ["-{:X}".format(i)  for i in range(16)] )
#         table.setVerticalHeaderLabels(   ["{:X}-".format(i) for i in range(16)] )
#         table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
#         table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
#         table.resizeColumnsToContents()


#     def timer_tick(self):
#         if self.clearColorCD > 0 :
#             self.clearColorCD -= 1
        
#         if self.clearColorCD == 0 :
#             self.clearColorCD = -1
#             c = self.Table_Rx.horizontalHeaderItem(0).background()
#             for row in range(16):
#                 for col in range(16):
#                     if self.Table_Rx.item(row, col) is not None :
#                         self.Table_Rx.item(row, col).setBackground(c)
#                     if self.Table_Tx.item(row, col) is not None :
#                         self.Table_Tx.item(row, col).setBackground(c)



#     def highlight_cell(self, wg:QTableWidget, address, datas, color):

#         self.clearColorCD = 5

#         for i, data in enumerate(datas)  :
#             row = (address + i)/16
#             col = (address + i)%16

#             item = QTableWidgetItem("{:02X}".format(data))
#             item.setTextAlignment(Qt.AlignCenter)
#             item.setBackground(QColor(color))
#             wg.setItem(row, col, item)        



    # -----------------------------------------------------------------------------------------------------------------
    # --  User Firing jobs (Threaading) -------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------

    @threadin_user_button
    @log_error
    def readClicked(self):
        device  = int(self.Text_DevAddr.text(), base=16)
        regaddr = int(self.Text_RegAddr.text(), base=16)
        len     = int(self.Text_ReadLen.text())

        self.read(device, regaddr, len)    

    @threadin_user_button
    @log_error
    def writeClicked(self):
        device  = int(self.Text_DevAddr.text(), base=16)
        regaddr = int(self.Text_RegAddr.text(), base=16)
        hexvals = [int(hexstr, base=16) for hexstr in re.split('\s+', self.Text_WriteData.toPlainText())]

        self.write(device, regaddr, hexvals)

    def clearCells(self):
        datas = [None] * 256
        self.RxModel.updateData(0x00, datas)
        self.TxModel.updateData(0x00, datas)

    # -----------------------------------------------------------------------------------------------------------------
    # --  EXPORT (thread-safe) ----------------------------------------------------------------------------------------
    # -----------------------------------------------------------------------------------------------------------------
    @run_in_main()
    def printscreen(self, savepath, frame = 'RX'):
        self.showForm()
        if frame  == 'RX' : 
            rect = self.GBox_Rx.frameGeometry()
            pixmap = QPixmap(rect.size()); 
            self.GBox_Rx.render(pixmap, QPoint(), QRegion(rect))
        else :
            rect = self.Table_Tx.frameGeometry()
            pixmap = QPixmap(rect.size()); 
            self.Table_Tx.render(pixmap, QPoint(), QRegion(rect))                
                
        pixmap.save(savepath)
        return pixmap


    @ezStatus    
    def read(self, device, regaddr, len):  
        """
        GI2C.read(device, regaddr, len) -> datas
        ================================================================
        [GI2C UI read i2c] 
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param len: length for read.
        :return: datas
        """
        datas = GJIGManager.MyJIG('GI2C').i2c_read(device, regaddr, len)

        self.setStatus(f"[0x{device:02X}],@{regaddr:02X}h,len={len} Read Done.")    
        self.RxModel.updateData(regaddr, datas)

        return datas

    @ezStatus  
    def write(self, device, regaddr, datas):  
        """        
        GI2C.write(device, regaddr, datas)                   
        ================================================================
        [GI2C UI Write i2c] 
        :param device: I2C slave address.
        :param regaddr: register address, command table address.
        :param datas: list for bytes to write.
        :return: None
        """
        GJIGManager.MyJIG('GI2C').i2c_write(device, regaddr, datas)

        self.setStatus(f"[0x{device:02X}],@{regaddr:02X}h,len={len(datas)} Write Done.")    
        self.TxModel.updateData(regaddr, datas)


    @ezStatus  
    def opread(self, device, len, contAck = True, lastAck = False):  
        """
        GI2C.opread(device, regaddr, len, contAck, lastAck) -> datas                 
        ================================================================
        [GI2C UI Optional-Read I2C] 
        :param device: I2C slave address.
        :param len: length for read.
        :param contAck: 1 for middle Ack. otherwise 0
        :param lastAck: 1 for last Ack.   otherwise 0
        :return: None
        """
        datas = GJIGManager.MyJIG('GI2C').i2c_opread(device, len, contAck, lastAck)

        self.RxModel.updateData(0xF0, datas)

        return datas


    @ezStatus  
    def opwrite(self, device, datas):  
        """
        GI2C.opwrite(device, datas): -> None
        ================================================================
        [GI2C UI Write i2c]
        :param device: I2C slave address.
        :param datas: list for bytes to write.
        :return: None
        """
        GJIGManager.MyJIG('GI2C').i2c_opwrite(device, datas)

        self.TxModel.updateData(0xF0, datas)




GI2C = I2C()
