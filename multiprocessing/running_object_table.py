"""
Example code to show how to use the 'Running Object Table' to find the right
Excel Application object to use from a child process.

When spawning a child process (or processes) from Excel sometimes it's
useful to be able to get a reference to the Excel Applciation object
corresponding to the parent process (to write back to a worksheet for
example).

The Excel application object can be obtained by using
win32com.client.Dispatch("Excel.Application") but this isn't guarenteed
to be the correct instance of Excel if there are multiple Excel applications
running.

This example shows how to solve this problem using the 'Running Object Table'.

References:
https://msdn.microsoft.com/en-us/library/windows/desktop/ms695276%28v=vs.85%29.aspx
https://msdn.microsoft.com/en-us/library/windows/desktop/ms684004(v=vs.85).aspx
"""
from pyxll import xl_func, get_active_object
from multiprocessing import Process
import win32com.client
import pythoncom
import logging
import time
import os


IID_Workbook = pythoncom.MakeIID("{000208DA-0000-0000-C000-000000000046}")


def get_xl_app(parent=None):
    """
    Return an Excel Application instance.

    Unlike using win32com.client.Dispatch("Excel.Application") the
    Application returned will always be the one that corresponds
    to the parent process.
    """
    # Get the window handle set by the parent process
    parent_hwnd = os.environ["PYXLL_EXCEL_HWND"]

    # Iterate over the running object table looking for the Excel Workbook
    # object from the parent process' Application object.
    context = pythoncom.CreateBindCtx(0)
    for moniker in pythoncom.GetRunningObjectTable():
        try:
            # Workbook implements IOleWindow so only consider objects implementing that
            window = moniker.BindToObject(context, None, pythoncom.IID_IOleWindow)
            disp = window.QueryInterface(pythoncom.IID_IDispatch)

            # Get a win32com Dispatch object from the PyIDispatch object as it's
            # easier to work with.
            obj = win32com.client.Dispatch(disp)

        except pythoncom.com_error:
            # Skip any objects we're not interested in
            continue

        # Check the object we've found is a Workbook
        if getattr(obj, "CLSID", None) == IID_Workbook:
            # Get the Application from the Workbook and if its window matches return it.
            xl_app = obj.Application
            if str(xl_app.Hwnd) == parent_hwnd:
                return xl_app

    # This can happen if the parent process has terminated without terminating
    # the child process.
    raise RuntimeError("Parent Excel application not found")


def _subprocess_func(target_address, logfile):
    """
    This function is run in a child process of the main Excel process.
    """
    # Initialize logging (since we're now running outside of Excel this isn't done for us).
    logging.basicConfig(filename=logfile, level=logging.INFO)
    log = logging.getLogger(__name__)
    log.info("Child process %d starting" % os.getpid())

    try:
        # Get the Excel Application corresponding to the parent process
        xl_app = get_xl_app()

        # Write to the target cell in the parent Excel.
        cell = xl_app.Range(target_address)
        message = "Child process %d is running..." % os.getpid()
        cell.Value = message

        # Run for a few seconds updating the value periodically
        for i in range(300):
            message = message[1:] + message[0]

            # When setting a value in Excel it may fail if the user is also
            # interacting with the sheet.
            try:
                cell.Value = message
            except:
                log.warn("Error setting cell value", exc_info=True)

            time.sleep(0.2)

        cell.Value = "Child process %d has terminated" % os.getpid()

    except Exception:
        log.error("An error occured in the child process", exc_info=True)
        raise


@xl_func("string target_address: string")
def start_subprocess(target_address):
    """
    Start a sub-process that will write back to a cell.
    :param target_address: address of cell to write to from the child process.
    """
    # Get the window handle of the Excel process so the sub-process can
    # find the right Excel Application instance.
    xl_app = win32com.client.Dispatch(get_active_object()).Application
    os.environ["PYXLL_EXCEL_HWND"] = str(xl_app.Hwnd)

    # Get the log file name for the subproces to log to.
    root = logging.getLogger()
    logfile = None
    for handler in root.handlers:
        if isinstance(handler, logging.FileHandler):
            logfile = handler.baseFilename
            break

    # Start the subprocess that will write back to the target cell.
    # It's a daemon process so that it doesn't stop the main Excel process
    # from terminating even if it's still running.
    process = Process(target=_subprocess_func, args=(target_address, logfile))
    process.daemon = True
    process.start()

    return "Child process %d started" % process.pid
