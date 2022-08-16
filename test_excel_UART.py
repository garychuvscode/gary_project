import pyvisa
# pyvisa is to used for GPIB and UART interface
import xlwings as xw
# xlwings is the API used to operating excel


print("before the UART is out")
# for the UART signal sub function
rm = pyvisa.ResourceManager()
# create new control object from pyvisa
UART = rm.open_resource("COM10")
# choose the related COM port for UART output
UART.write("abcde")
# the content need to output
print("after the UART is out")


# app = xw.App(visible=True, add_book=True)
# open excel program and open new workbook
# wb = app.books.add()
# open new workbook named wb
app = xw.Book('testing name.xlsx')
# app.Range('A1') = 1
# app.Range('A2:C2') = 2
# app.Range((6, 3)) = 3
