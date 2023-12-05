# turn off the formatter
# fmt: off

try:
    # run the try items
    print(f'now is running try items')

except Exception as e:
    # show the exception but not stop the program
    # e is the exception message from program
    print(f"there are error {str(e)}")

    # if you want to stop the program with exception show
    raise Exception(
        f'<>< Chamber_SU242 ><> open Chamber Fail {str(e)}!')
