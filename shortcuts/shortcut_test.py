"""
Test code from the docstring in shortcuts.py.

Add this to your pyxll.cfg to test.
"""
from pyxll import xl_macro, xlcAlert
from shortcuts import xl_shortcut

@xl_shortcut("Ctrl+Shift+H")
@xl_macro("")
def my_macro():
    xlcAlert("Hello!")
