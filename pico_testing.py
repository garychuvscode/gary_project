# 假设有一个包含多个小于等于255的数字的列表
x_values = [100, 150, 200]

# 将列表直接转换为bytes对象
data_to_write = bytes(x_values)

# 打印结果
print("Data to Write:", data_to_write)

numeric_list = [byte for byte in data_to_write]

print(f"Data to Write, number:{numeric_list}")
