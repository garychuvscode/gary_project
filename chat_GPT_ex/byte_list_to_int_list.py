import struct

# 原始的bytes
byte_data = b"\x00\x0c\x00\x01"

# 使用struct.unpack转换为int list
int_list = struct.unpack("<BBBB", byte_data)
# 这样的结果应该是 [0, 12, 0, 1]。这个格式字符串 <BBBB 表示按照小端字节顺序解包为四个无符号字节。

print(int_list)
