"""
如果你想将找到的值存入类的实例变量（例如 self.ind_row、self.ind_col),
你可以通过直接使用 setattr 函数来设置实例变量。以下是相应的修改：
"""


class Example:
    def __init__(self):
        # 初始化变量
        self.ind_row = None
        self.ind_col = None
        pass

    def process_kwargs(self, **kwargs):
        # 定义要匹配的关键字
        valid_keys = {"name", "age", "city", "profession"}

        # 遍历 kwargs 中的键值对
        for key, value in kwargs.items():
            # 如果键在 valid_keys 中，将对应的值存入相应的实例变量中
            if key in valid_keys:
                setattr(self, f"ind_{key}", value)

        # 打印结果
        print(f"Name: {self.ind_name}")
        print(f"Age: {self.ind_age}")
        print(f"City: {self.ind_city}")
        print(f"Profession: {self.ind_profession}")


# 创建 Example 类的实例
example_instance = Example()

# 调用实例方法
example_instance.process_kwargs(
    name="John", age=25, city="New York", profession="Engineer"
)
