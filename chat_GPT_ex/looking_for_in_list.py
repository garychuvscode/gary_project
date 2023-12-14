# turn off the formatter
# fmt: off

# 示例列表
my_list = ["6V", "9V", "12V", "18V", "24V"]

# 要搜索的值
search_value = "12V"

# 使用 index 方法查找值的索引
index_of_value = my_list.index(search_value)

# 输出结果
print(f"The index of {search_value} in the list is: {index_of_value}")


# method to prevent issue of searching, search first

if search_value in my_list:
    index_of_value = my_list.index(search_value)
    print(f"The index of {search_value} in the list is: {index_of_value}")
else:
    print(f"{search_value} is not in the list.")


# another example

my_list = ["a", "b", "c"]

try:
    index_of_b = my_list.index("b")
    print(f"The index of 'b' is: {index_of_b}")
except ValueError:
    print("Element not found in the list.")
