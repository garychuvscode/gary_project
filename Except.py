
# -----------------------------------------------------------------------------------------------------------------
# --  JIGError  ---------------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------   
class JIGError(Exception):
    pass


# -----------------------------------------------------------------------------------------------------------------
# --  JIGError\I2CError  ------------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------------------------------   
class I2CError(JIGError):
    pass

class I2CBusError(I2CError):
    pass

class I2CNACKError(I2CError):
    pass