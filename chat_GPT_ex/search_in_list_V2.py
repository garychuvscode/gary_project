# turn off the formatter
# fmt: off

test_index = 4

if test_index == 0 :

    # 示例列表
    my_string_list = ["abc12xyz", "def", "123", "gh12ij", "klm"]

    # 搜索的子字符串
    search_substring = '12'

    # 使用列表推导式筛选包含子字符串 '12' 的元素
    result_elements = [element for element in my_string_list if search_substring in element]

    # 输出结果
    print(f"The elements containing '{search_substring}' are: {result_elements}")

if test_index == 1:
    # to record the index

    # 示例列表
    my_string_list = ["abc12xyz", "def", "123", "gh12ij", "klm"]

    # 搜索的子字符串
    search_substring = '12'

    # 用于存储第一个包含子字符串 '12' 的元素的索引
    result_index = None

    # 遍历列表，记录第一个包含子字符串 '12' 的元素的索引
    for index, element in enumerate(my_string_list):
        if search_substring in element:
            result_index = index

            break

    # 输出结果
    print(f"The first index of an element containing '{search_substring}' is: {result_index}")

if test_index == 2:
    # to record the index

    # 示例列表
    my_string_list = ["abc12xyz", "def", "123", "gh12ij", "klm"]

    # 搜索的子字符串
    search_substring = '12'

    # 用于存储包含子字符串 '12' 的元素的索引
    result_indices = []

    # 遍历列表，记录包含子字符串 '12' 的元素的索引
    for index, element in enumerate(my_string_list):
        if search_substring in element:
            result_indices.append(index)

    # 输出结果
    print(f"The indices of elements containing '{search_substring}' are: {result_indices}")

if test_index == 3:
    # include both index and content

    # 示例列表
    my_string_list = ["abc12xyz", "def", "123", "gh12ij", "klm"]

    # 搜索的子字符串
    search_substring = '12'

    # 用于存储包含子字符串 '12' 的元素和它们的索引的列表
    result_elements_and_indices = []

    # 遍历列表，记录包含子字符串 '12' 的元素和它们的索引
    for index, element in enumerate(my_string_list):
        if search_substring in element:
            # this use tuple as the output result
            result_elements_and_indices.append((element, index))

    # 输出结果
    print(f"The elements and their indices containing '{search_substring}' are: {result_elements_and_indices}")

if test_index == 4:
    # 示例列表
    my_string_list = ["abc12xyz", "def", "123", "gh12ij", "klm"]

    # 搜索的子字符串
    search_substring = '12'

    # 用于存储包含子字符串 '12' 的元素的列表
    result_elements = []

    # 用于存储包含子字符串 '12' 的元素的索引的列表
    result_indices = []

    # 遍历列表，记录包含子字符串 '12' 的元素和它们的索引
    for index, element in enumerate(my_string_list):
        '''
        在这个例子中,enumerate 函数用于同时获取元素和它们的索引。然后，
        使用列表推导式检查每个元素是否包含子字符串
        '12'，并返回包含该子字符串的元素的索引列表。
        '''
        if search_substring in element:
            result_elements.append(element)
            result_indices.append(index)

    # 输出结果
    print(f"The elements containing '{search_substring}' are: {result_elements}")
    print(f"The indices of elements containing '{search_substring}' are: {result_indices}")
