
import pyxll
from pyxll import xl_func


@xl_func("int x, int x: int")
def addTwoNumbersByDavidKe(x, y):
    """returns the sum of a range of floats"""
    return x + y

@xl_func("int x: int")
def echoByDavidKe(x):
    """ """
    return x


