def process_kwargs(**kwargs):
    # 定义要匹配的关键字
    valid_keys = {"name", "age", "city", "profession"}

    # 初始化变量
    name = age = city = profession = None

    # 遍历 kwargs 中的键值对
    for key, value in kwargs.items():
        # 如果键在 valid_keys 中，将对应的值存入相应的变量中
        if key in valid_keys:
            locals()[key] = value

    # 打印结果
    print(f"Name: {name}")
    print(f"Age: {age}")
    print(f"City: {city}")
    print(f"Profession: {profession}")


# 调用函数
process_kwargs(name="John", age=25, city="New York", profession="Engineer")
