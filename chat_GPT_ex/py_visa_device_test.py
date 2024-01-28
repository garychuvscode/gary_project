import pyvisa
import time

rm = pyvisa.ResourceManager()
available_devices = rm.list_resources()

com_addr = 6
uart_cmd_str = f"COM{com_addr}"

for device in available_devices:
    print(device)

    try:
        # to check what is pico
        pico = rm.open_resource(uart_cmd_str)
        pico.clear()

        dev_name = 0
        pico.write("*IDN?")
        # the first of read in pico after write is to get command
        cmd_write = pico.read()
        # the second read is the return item (if there are return)
        item_back = pico.read()
        print(item_back)

        print(
            f"what we got on usb is: first the command {cmd_write},second the item_back {item_back}"
        )

    except Exception as e:
        # may not have or wrong device
        print(f"exception: {e}, please check pico connection")


# pico = rm.open_resource("ASRL6::INSTR")
a = pico.query("t;usb")
print(a)

a = pico.query("t;i", 5)
# the query will
print(a)

b = pico.read()
print(b)

c = pico.read()
print(c)


# time.sleep(10)
