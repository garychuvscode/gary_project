# turn off the formatter
# fmt: off

def multiple_types_list():
    return [1, 'hello', {'key': 'value'}]

result_list = multiple_types_list()
print(result_list)  # 输出 [1, 'hello', {'key': 'value'}]


def multiple_types():
    return 1, 'hello', [1, 2, 3]

result = multiple_types()
print(result)  # 输出 (1, 'hello', [1, 2, 3])


def multiple_types_dict():
    return {'integer': 1, 'string': 'hello', 'list': [1, 2, 3]}

result_dict = multiple_types_dict()
print(result_dict)  # 输出 {'integer': 1, 'string': 'hello', 'list': [1, 2, 3]}
