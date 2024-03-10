def extract_specific_bit(hex_list, byte_index, bit_index):
    # 檢查位元組索引是否在範圍內
    if byte_index < 0 or byte_index >= len(hex_list):
        return None

    # 將位元組轉換為二進位字符串
    byte_binary = bin(int(hex_list[byte_index], 16))[2:].zfill(8)

    # 檢查位元索引是否在範圍內
    if bit_index < 0 or bit_index >= 8:
        return None

    # 返回特定位元的值
    return byte_binary[bit_index]


# 使用示例
hex_list = ["FF", "00", "A5"]  # 假設 hex_list 是一個十六進位數字列表
byte_index = 0  # 取第一個位元組
bit_index = 3  # 取第四個位元
result = extract_specific_bit(hex_list, byte_index, bit_index)
print(f"the bit is: {result}")  # 輸出結果
