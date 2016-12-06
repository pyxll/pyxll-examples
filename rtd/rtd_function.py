"""
** For an easier way to do real time data in Excel with Python **
** see https://www.pyxll.com/docs/userguide/rtd.html **

This example shows how to write an Excel function using PyXLL
that connect to an RTD server and return values from it.

When the RTD server publishes updates Excel will automatically update
the cell the RTD function is called from.

The RTD server from the `exceltypes` package is used as the example
server and so that will need to be installed before this demo will work.

The `exceltypes` package can be downloaded from here:
https://github.com/pyxll/exceltypes
"""
from pyxll import xl_func, get_active_object
import win32com.client
import win32api
import win32con
import pythoncom
import win32com.client

# the TimeServer from the exceltypes demos is used as the RTD server
from exceltypes.demos.excelRTDServer import TimeServer


# calling this function from Excel returns a value from the RTD server that ticks
# when the RTD server publishes updates.
@xl_func("int seconds: var")
def py_rtd_test(seconds):
    xl = win32com.client.Dispatch(get_active_object()).Application
    return xl.WorksheetFunction.RTD(TimeServer._reg_progid_, "", "seconds", str(seconds))


def _register(cls):
    """Register an inproc com server in HKEY_CURRENT_USER.

    This may be used as a replacement for win32com.server.register.UseCommandLine
    to register the server into the HKEY_CURRENT_USER area of the registry
    instead of HKEY_LOCAL_MACHINE.
    """
    clsid_path = "Software\\Classes\\CLSID\\" + cls._reg_clsid_
    progid_path = "Software\\Classes\\" + cls._reg_progid_
    spec = cls.__module__ + "." + cls.__name__

    # register the class information
    win32api.RegSetValue(win32con.HKEY_CURRENT_USER, clsid_path, win32con.REG_SZ, cls._reg_desc_)
    win32api.RegSetValue(win32con.HKEY_CURRENT_USER, clsid_path + "\\ProgID", win32con.REG_SZ, cls._reg_progid_)
    win32api.RegSetValue(win32con.HKEY_CURRENT_USER, clsid_path + "\\PythonCOM", win32con.REG_SZ, spec)
    hkey = win32api.RegCreateKey(win32con.HKEY_CURRENT_USER, clsid_path + "\\InprocServer32")
    win32api.RegSetValueEx(hkey, None, None, win32con.REG_SZ, pythoncom.__file__)
    win32api.RegSetValueEx(hkey, "ThreadingModel", None, win32con.REG_SZ, "Both")

    # and add the progid
    win32api.RegSetValue(win32con.HKEY_CURRENT_USER, progid_path, win32con.REG_SZ, cls._reg_desc_)
    win32api.RegSetValue(win32con.HKEY_CURRENT_USER, progid_path + "\\CLSID", win32con.REG_SZ, cls._reg_clsid_)

# make sure the example RTD server is registered
_register(TimeServer)
