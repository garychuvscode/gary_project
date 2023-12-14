class Example:
    pass


# 创建 Example 类的实例
example_instance = Example()

# 使用 setattr 设置对象属性
setattr(example_instance, "name", "John")
setattr(example_instance, "age", 25)

# 访问对象属性
print(example_instance.name)  # 输出: John
print(example_instance.age)  # 输出: 25
