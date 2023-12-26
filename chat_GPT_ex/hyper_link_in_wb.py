import xlwings as xw

# 打開或創建一個 Excel 工作簿
wb = xw.Book()

# 獲取活動工作表
sheet = wb.sheets.active

# 在 A1 儲存格輸入文字
sheet.range("A1").value = "Click here!"

# 設定超連結到 B2 儲存格
link_address = sheet.range("B2").address
sheet.range("A1").api.Hyperlinks.Add(
    Anchor=sheet.range("A1").api,
    Address="",  # 留空，表示連結到同一個工作簿
    SubAddress=f"{sheet.name}!{link_address}",
)

# 儲存工作簿
wb.save("example.xlsx")

# 顯示 Excel
wb.app.visible = True
