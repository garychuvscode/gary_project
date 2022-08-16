
import pyvisa
from win32con import ITALIC_FONTTYPE

rm = pyvisa.ResourceManager()
rm.list_resources()
#('ASRL1::INSTR', 'ASRL2::INSTR', 'GPIB0::14::INSTR')
#my_instrument = rm.open_resource('GPIB0::14::INSTR')
# print(my_instrument.query('*IDN?'))

COM_test = 0
# COM_test: 0=> is the testing on MCU
# 1 => is the testing for different content

# to send related command to comport, need to use array to map
array = ['01', 'FF', '04', '08', '10', '20', '40', '80']

print(array[1])
print(array[2])
print(array[3])
print(array[4])
print(array[5])


test = hex(255)
print(test)
test = str(test)
print(test)


com_num = 7
mode = 5

itc4 = rm.open_resource("COM" + str(com_num))
print("COM" + str(com_num))

if COM_test == 0:
    for i in range(0, 8):
        # control variable can map to the output directly
        # by using 1-8 for related GPIO

        itc4.write(chr(mode) + array[i])
        print(array[i])
        input()


if COM_test == 1:

    itc4.write("abcde")
    # print(itc4.read())
    back = ' not yet '
    # COM port testing

    # first is the testing of mode
    cmd1 = 1
    itc4.write(str(int(cmd1)))
    # back = itc4.read()
    print('first cmd ' + str(cmd1))
    print('the reply ' + str(back))
    temp = input()

    # second is the testing of string (I2C command)
    cmd1 = 'FF'
    itc4.write(str(cmd1))
    # back = itc4.read()
    print('second cmd ' + str(cmd1))
    print('the reply ' + str(back))
    temp = input()

    # third is the testing of the SWIRE
    cmd1 = 70
    itc4.write(str(int(cmd1)))
    # back = itc4.read()
    print('third cmd ' + str(cmd1))
    print('the reply ' + str(back))
    temp = input()

    # fourth is the testing of GPIO (GPIO control)
    cmd1 = '1'
    itc4.write(str(cmd1))
    # back = itc4.read()
    print('fourth cmd ' + str(cmd1))
    print('the reply ' + str(back))
    temp = input()

    # fifth is the testing of GPIO (GPIO control), SWIRE
    cmd1 = 400
    # print(chr(cmd1))
    x = chr(cmd1)
    # chr function is the change int to the char, make transmit can be only number
    # no need to build up conversion at the MCU side
    print(chr(cmd1))

    # "itc4.write(chr(cmd1))" this command will occur error because
    # write function of COM port can only transmit number lower than 128
    # there are no issue at the chr() function, it can transfer much bigger than 128

    # for application of SWIRE, need to use times to send , maximum 128 * 2
    # should be enough

    # back = itc4.read()
    print('fourth cmd ' + str(cmd1))
    print('the reply ' + str(back))

    temp = input()

    print('testing,123')
    # This program prints Hello, world!

    print('Hello, world!123')

    # input and output example
    # a = input()
    # print('what you key in is:' + a)

    # This program adds two numbers

    num1 = 1.5
    num2 = 6.3

    # Add two numbers
    sum = num1 + num2

    # Display the sum
    print('The sum of {0} and {1} is {2}'.format(num1, num2, sum))
