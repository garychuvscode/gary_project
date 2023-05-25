# excel related package

import win32api
from win32con import MB_SYSTEMMODAL

# from pickle import NONE
import xlwings as xw

print("testing output ")
print("testing output123 ")

# turn off the formatter
# fmt: off

# 匯入 xlwings 套件
# 讓 xlwings 動態開啟一個新的 Excel 檔案，並將該檔案存入 workbook 變數
# wb = xw.App()
# app() 內部function較少不使用
wb = xw.Book('C:\\py_gary\\test_excel\\new.xlsx')
wb = xw.Book()
wb2 = xw.Book()

wb2.sheets.add('test')
wb2.sheets
wb2.save('c:\\py_gary\\test_excel\\test.xlsx')
# file path in python need to be using \\ , not only one \;
# need to use full name and trace(change name here)

sheet1 = wb2.sheets("test")

# sheet2 = wb2.sheets('工作表1')
# delete the default sheet
wb2.sheets('工作表1').delete()
# sheet2.delete()
sheet1.range('A1:B3').value = 'python知识学堂'
sheet1.range('A1').clear()
sheet1.range('G1').value = [['a', 'b', 'c'], [1, 2, 3]]

sheet1.range((10, 10), (3, 10),).value = '1010abc'
# multi cell value assign

# single cell value assign
sheet1.range((9, 1)).value = "A9"

xls_app = xw.App(visible=True, add_book=False)
xls_app.display_alerts = False
# turn off the displat alert (asking to remove sheet...)

sheet1.copy()
sheet1.delete()
# separate input data in different row


wb2.save('c:\\py_gary\\test_excel\\test.xlsx')
# wb2.close()

# close the workbooks


# xw.sheets.add('test_sheet')
# 若你沒有開啟你的 Excel, 這行可能會跑的慢一點，請耐心等待，讓子彈飛一會兒～

# ===== first way to implement class, object application
# application of class and obect
class Human:
    def introduce(self):
        print('my name is:' + self.name)

    def stand(self):
        print(self.name + ' stands up')
# finished define the class


# start to initialization for the variable
h1 = Human()
h1.name = 'Tom'
h1.weight = 70
h1.language = 'english'

# call the method in Human object
h1.introduce()

# second way to create class and object
# define with constructor


class Dog:
    def __init__(self, name_d, color_d, weight_d, owner_d):
        self.name = name_d
        self.color = color_d
        self.weighr = weight_d
        self.owner = owner_d

    def bark(self):
        print(self.name + ' is barking')

    def sleep(self):
        print(self.name + ' is sleeping')

# application with constructor


d1 = Dog('bob', 'black', 70, None)
d2 = Dog('happy', 'white', 50, None)

d1.bark()
d2.sleep()
h1.stand()

# define the onwer of the dog
d2.owner = h1
print('the owner of ' + d2.name + ' is ' + d2.owner.name)
# interatcion between different object
# connection build by different variable

# === for loop demonstration

# from 0-9 => < 10 and start from 0 (like array)
for i in range(10):
    print(i, end=' ')

print('')
# change to next row

# for loop similiar with C
# (start, end(<=), step)
for i in range(4, 30, 2):
    print(i, end=' ')
    # end=' ' => means not change toi next row

print('')

# (start, end, step(default=1))
for i in range(4, 30):
    print(i, end=' ')

print('')

sum = 0
i = 1
while i <= 10:
    sum = sum + i
    i = i + 1
print(sum, end=" ")

print('')

for i in range(1, 10):
    for j in range(1, 10):
        if j == 9:
            print("\t", i*j)  # j == 9時，換行
        else:
            print("\t", i*j, end='')  # j < 9時，不換行

print('')

for i in "Hey Jude":
    if i == "u":
        break
    print(i)

print('')

for i in "Hey Jude":
    if i == "u":
        continue
    print(i)

print('')

i = 1
while i <= 10:
    print(i, end=" ")
    i = i + 1

print('')

sequences = [0, 1, 2, 3, 4, 5]
i = 0
while 1:  # 判斷條件值為1, 代表迴圈永遠成立
    print(sequences[i], end=" ")
    i = i + 1
    if i == len(sequences):
        print("No elements left.")
        break

print('')


response = win32api.MessageBox(
    0, "Did you hear the Buzzer?", "Buzzer Test", 4, MB_SYSTEMMODAL)

print('')

win32api.MessageBox(0, 'hello', 'title')

print('')

# subfunction test

ch_num0 = 1
v_res0 = 1


def measure_v(ch_num, v_res):
    print('we have below input variable: ' + str(ch_num) + 'and ' + str(v_res))
    # by using sub to measure and get result into v_res
    # call GPIB measure function and
    if v_res == 1:
        v_res = 0
    else:
        v_res = 1
    print('measurement result is ' + str(v_res))

    # after the measure is finished change to another channel
    if ch_num > 6:
        # channel number can be change
        ch_num = 1
        print('channel reset')
    else:
        ch_num = ch_num + 1
        print('channel increase')
    return v_res, ch_num
    # return can follow the sequence, not related to the sequence of input variable


v_res0, ch_num0 = measure_v(ch_num0, v_res0)

v_res0, ch_num0 = measure_v(ch_num0, v_res0)

v_res0, ch_num0 = measure_v(ch_num0, v_res0)

print('testing the global variable')

# the other way of implement global variable change is to use below
# define the global variable: global name


def measure_v2(ch_num, v_res):
    global ch_num0
    global v_res0
    print('we have below input variable: ' + str(ch_num) + 'and ' + str(v_res))
    # by using sub to measure and get result into v_res
    # call GPIB measure function and
    if v_res == 1:
        v_res = 0
    else:
        v_res = 1
    print('measurement result is ' + str(v_res))

    # after the measure is finished change to another channel
    if ch_num > 6:
        # channel number can be change
        ch_num = 1
        print('channel reset')
    else:
        ch_num = ch_num + 1
        print('channel increase')
    # take off the return and using the change of global variable
    # return v_res, ch_num
    # return can follow the sequence, not related to the sequence of input variable

    # update the global variable at the end of the sub program
    ch_num0 = ch_num
    v_res0 = v_res


measure_v2(ch_num0, v_res0)

measure_v2(ch_num0, v_res0)

measure_v2(ch_num0, v_res0)

measure_v2(ch_num0, v_res0)


print('')

# test for object to change global variable
# if this will cause error

dog_chg = 10


class Dog2:
    def __init__(self, name_d, color_d, weight_d, owner_d):
        self.name = name_d
        self.color = color_d
        self.weighr = weight_d
        self.owner = owner_d

    def bark(self):
        print(self.name + ' is barking')

    def sleep(self):
        print(self.name + ' is sleeping')

    def change(self):
        print('try to change the global variable dog_chg ')
        global dog_chg
        dog_chg = 5


print(dog_chg)

# application with constructor
e1 = Dog2('age ', 'black', 70, None)
e2 = Dog2('sad ', 'white', 50, None)

e1.bark()
e2.sleep()


e1.change()
print(dog_chg)
# these prove that method in object created by class, can also change global function


print('')

# string replacement practice
test_str = 'this is a test string'
print(test_str)
test_str = test_str.replace('is', 'ab')
print(test_str)


print('')
print('')
