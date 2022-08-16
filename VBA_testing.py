
import win32api
import win32com.client
import win32con
import os
import time
import xlwings as xw

# ======== xlwings related part
control_book_trace = 'c:\\py_gary\\test_excel\\eff_chart_test.xlsm'
# no place to load the trace from excel or program, define by default
result_book_trace = ''
# result trace unable to load yet
# wb = xw.Book(control_book_trace)
# open from the trace and assign to the object
# ======== xlwings related part


test_step = 3
# use to control the testing program

# ======== excel application related
# 開啟 Excel 的app
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
# ======== excel application related

if test_step == 0:
    # 單純使用 add function 來判斷 active sheet 的影響, add_one1 的function 才可以正確執行,
    # add_one如果active sheet change, 就會出現error
    # 開啟 ctrl_test.xlsm 活頁簿檔案
    excel.Workbooks.Open(Filename="C:\\py_gary\\test_excel\\ctrl_test.xlsm")

    # 執行巨集程式
    # 產生新的file
    excel.Application.Run("ctrl_test.xlsm!new_books",
                          "test_gen", "ctrl_test.xlsm")

    x_add = 0
    while x_add < 20:
        excel.Application.Run("ctrl_test.xlsm!add_one", 5, 3)
        excel.Application.Run("ctrl_test.xlsm!add_one1",
                              2, 2, 'test_gen.xlsx', '工作表1')
        # if there are no sheet information input to the sub program, it will be operated to sheet active
        # change the window during operation will cause problem of saving the data
        # must have sheet information when calling VBA: specific sheet(books), and x-axis, y-axis
        time.sleep(1)
        print(x_add)
        x_add = x_add + 1

    # filename!subname,

    # 傳入參數，並取得計算結果
    # result = excel.Application.Run("test_gen", "0823_SWIRE_scan_IQ")
    # print(result)

elif test_step == 1:
    # testing for plot fnction without assign the workbook and worksheet
    # 此方法為data 在 new1 裡面, 但如果control sheet 裡面才儲存巨集, 目標data會在不同的workbook 裡面

    # 先用excel app 開啟後才assign 到 xlwings

    # 開啟 hello.xlsm 活頁簿檔案
    excel.Workbooks.Open(Filename="C:\\py_gary\\test_excel\\ctrl_test.xlsm")

    wb = xw.books('ctrl_test.xlsm')
    # assign the book already open to the xlwings opject
    # wb = xw.Book(control_book_trace)
    # open new book from trace

    v_cnt = 27
    i_cnt = 31
    sheet_n = 'Efficiency'
    excel.Application.Run("ctrl_test.xlsm!gary_chart", v_cnt, i_cnt, sheet_n)

    sheet_n = 'load regulation_ELVDD'
    excel.Application.Run("ctrl_test.xlsm!gary_chart", v_cnt, i_cnt, sheet_n)

    # calling the save function from xlwings
    wb.save()
    # 利用xlwings 來做儲存, 後面才可直接關閉


elif test_step == 2:
    # open file xlwings and assign the book already open to the excel app
    # this should be the bset way to fit current system
    # since the main operation is using xlwings, better to join in assign, not change the open file method
    # 因為目前大多都是xlwings 所以盡量不要更改開檔方式, 看是否可以將book and sheet assign 給excel app 呼叫巨集
    # 也可讓excel app 重新開啟control sheet, xlwings 後面住要reference to wb_res 所以control sheet 可以關閉後重新開啟
    # (open from excel app)就可以直接呼叫 => double check if sheet name can still used
    result_book_trace = 'c:\\py_gary\\test_excel\\eff_raw_test.xlsx'
    # wb = xw.Book(result_book_trace)
    # open from the trace and assign to the object

    # time.sleep(1)
    excel.Workbooks.Open(
        Filename="C:\\py_gary\\test_excel\\eff_chart_test.xlsm")
    # use excel app to open the control book, so we can call VBA function
    # time.sleep(1)

    # use excel app to open the control sheet
    wb = xw.Book(result_book_trace)

    v_cnt = 27
    i_cnt = 31
    sheet_n = 'Efficiency'
    book_n = 'eff_raw_test.xlsx'
    excel.Application.Run(
        "eff_chart_test.xlsm!gary_chart", v_cnt, i_cnt, sheet_n, book_n)

    sheet_n = 'load regulation_ELVDD'
    excel.Application.Run(
        "eff_chart_test.xlsm!gary_chart", v_cnt, i_cnt, sheet_n, book_n)

    wb.save()
    # excel.Application.Quit()

if test_step == 100:

    excel.Application.Quit()
    del excel
# 離開 Excel
# excel.Application.Quit()
# 會直接完全離開excel, close all the windows, need to choose save or not
else:
    # 清理 com 介面
    del excel
    # 清理介面之後才不會有檔案使用中的問題
