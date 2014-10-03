from pyxll import xl_macro
import win32api

@xl_macro()
def popup_messagebox():
    """pops up a message box"""
    win32api.MessageBox(0, "Hello", "Hello")

@xl_macro("string x: int")
def py_strlen(n):
    """returns the length of x"""
    return len(x)