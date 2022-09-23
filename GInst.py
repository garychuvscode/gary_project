# from PySide6.QtCore import *  # type: ignore
# from PySide6.QtGui import *  # type: ignore
# from PySide6.QtWidgets import *  # type: ignore

from functools import wraps

# original need to put "QObject" in the GInst class


class GInst():
    # valueUpdated  = Signal(float, str)
    # value2Updated = Signal(float, str)
    # onUpdated     = Signal()
    # offUpdated    = Signal()
    pass


def GInstSetMethod(unit='', value2=False):

    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # if value2 :
            #     args[0].value2Updated.emit(args[1], unit)
            # else :
            #     args[0].valueUpdated.emit(args[1], unit)

            result = func(*args, **kwargs)
            return result

        return wrapper
    return decorator


def GInstGetMethod(unit='', value2=False):

    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            result = func(*args, **kwargs)

            # if value2 :
            #     args[0].value2Updated.emit(result, unit)
            # else :
            #     args[0].valueUpdated.emit(result, unit)

            return result

        return wrapper
    return decorator


def GInstOnMethod():
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # args[0].onUpdated.emit()
            return func(*args, **kwargs)

        return wrapper
    return decorator


def GInstOffMethod():
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # args[0].offUpdated.emit()
            return func(*args, **kwargs)

        return wrapper
    return decorator
