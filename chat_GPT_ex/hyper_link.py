import xlwings as xw

# turn off the formatter
# fmt: off

# 開啟 Excel 應用程式
app = xw.App(visible=True)

# 開啟或建立工作簿
wb = app.books.add()

# 在第一個工作表的 A1 儲存格輸入內容
wb.sheets[0].range("A1").value = "Start Here"

# 在第二個工作表的 A1 儲存格輸入內容
wb.sheets.add("Sheet2").range("A1").value = "Target Cell"

# 在第二個工作表的 A1 儲存格添加超連結到第一個工作表的 A1 儲存格
target_sheet = wb.sheets[1]
target_sheet.range("A1").add_hyperlink(
    address=wb.sheets[0].name + "!A1", text_to_display="Go to Start", screen_tip=""
)

# 保存工作簿
wb.save(r"C:\py_gary\test_excel\hyperlink_example.xlsx")

# 關閉 Excel 應用程式
app.quit()

'''

add_hyperlink(address, text_to_display=None, screen_tip=None)
Adds a hyperlink to the specified Range (single Cell)

PARAMETERS
address (str) - The address of the hyperlink.

text_to_display (str, default None) - The text to be displayed for the hyperlink. Defaults to the hyperlink address.

screen_tip (str, default None) - The screen tip to be displayed when the mouse pointer is paused over the hyperlink. Default is set to ‘<address> - Click once to follow. Click and hold to select this cell.’

New in version 0.3.0.

'''
