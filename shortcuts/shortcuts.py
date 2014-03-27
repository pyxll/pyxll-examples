"""
Example code showing how PyXLL and win32com can be used to register macros and
associate them to keyboard shortcuts.

The decorator @xl_shortcut is used to decorate a macro function, and when
Excel starts win32com is used to associate that macro to a shortcut.

Example::

    from pyxll import xl_macro, xlcAlert
    from shortcuts import xl_shortcut

    @xl_shortcut("Ctrl+Shift+H")
    @xl_macro()
    def my_macro():
        xlcAlert("Hello!")

Pressing Ctrl+Shift+H calls the macro and shows the alert.

"""
from pyxll import get_active_object
from win32com.client import Dispatch
import logging
import timer
import time

_log = logging.getLogger(__name__)

# the shortcuts are added in the main window loop using timer
_timer_id = None
_shortcuts_to_add = []


def xl_shortcut(shortcut):
    """
    A decorator that can be applied to a PyXLL macro function
    that assocaiates a keyboard shortcut with the macro.
    
    The macro should take no arguments.
    
    :param shortcut: keyboard shortcut to assign to the macro,
                     e.g. Ctrl+Shift+R.
    :return: decorator to be applied to a macro function.
    """
    def make_decorator(shortcut):
        # convert the shortcut into a key code OnKey understands
        special_keys = {
            "ctrl"  : "^",
            "alt"   : "%",
            "shift" : "+",
        }

        accelerator = ""
        for key in shortcut.lower().split("+"):
            if key in special_keys:
                accelerator += special_keys[key]
                continue

            if key[0] == "f" and key[1:].isdigit():
                acc += "{F%s}" % key[1:]
                continue
            
            if len(key) > 1:
                raise Exception("Unrecognized shortcut %s" % shortcut)
            accelerator += key

        # the decorator that will add the shortcut and macroname to the list of
        # pending shortcuts to be added later.
        def xl_shortcut_decorator(func):
            global _timer_id

            # add the function and accelerator to the list of shortcuts to be processed
            _log.debug("Adding shortcut %s -> %s" % (shortcut, func.__name__))
            _shortcuts_to_add.append((accelerator, func.__name__))

            # start the timer to process the shortcuts list if necessary
            if _timer_id is None:
                _start_timer()

        return xl_shortcut_decorator

    return make_decorator(shortcut)


def _start_timer(timeout=300, interval=0.1):
    """
    Creates and starts a timer function that adds any registered shortcuts
    to Excel.

    As the COM object needed to do that may not exist at startup this
    function retries a number of times.
    """
    global _timer_id

    def make_timer_func():
        start_time = time.time()
        def on_timer(timer_id, unused):
            # when Excel is starting up and PyXLL imports its modules there
            # is no Excel window, and so the COM object needed to register
            # the shortcut might not be available yet.
            try:
                xl_window = get_active_object()
            except RuntimeError:
                if time.time() - start_time > timeout:
                    _log.error("Timed out waiting for Excel COM interface to become available.")
                    timer.kill_timer(timer_id)
                    _timer_id = None
                return

            # Get the Excel application from its window
            xl_app = Dispatch(xl_window).Application

            while _shortcuts_to_add:
                accelerator, macroname = _shortcuts_to_add.pop()
                try:
                    xl_app.OnKey(accelerator, macroname)
                except:
                    _log.error("Failed to add shortcut %s -> %s" % 
                                    (accelerator, macroname), exc_info=True)
            _log.debug("Finished adding shortcuts")
            timer.kill_timer(timer_id)
            _timer_id = None
        return on_timer

    # start the timer
    _timer_id = timer.set_timer(int(interval * 1000), make_timer_func())
