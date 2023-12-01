# 原始字符串
original_string = "i2c;w;addr;[1,2,3,4,5]"

# 使用分號分割字符串
split_parts = original_string.split(";")

# 提取包含整數的字符串，即 '[1,2,3,4,5]'
integer_string = split_parts[-1]

# 使用 eval 函數將字符串轉換為列表
integer_list = eval(integer_string)

integer_list_b = bytes(integer_list)
# 231128 this shuld be able to used in the I2C transfer

# 打印結果
print(integer_list)
