"""
Optimisation example using scipy.optimize.minimize.

Extended to ad.

This code accompanies the blog post
  https://www.pyxll.com/blog/XXXXXXXXXXXXX/
"""
import sys

from functools import partial
from pyxll import xl_macro, xl_app, xl_menu
from win32com.client import constants
import numpy as np
from scipy.optimize import minimize
from PyQt5.QtWidgets import QApplication, QMessageBox, QInputDialog, QPushButton

app = QApplication([sys.executable])

@xl_menu("sicpy.optimize")
def optimize4():
    xl = xl_app()
    # Get the initial values of the input cells
    in_values = list(xl.Range('Inputs').Value)
    X = np.array([x[0] for x in in_values])

    orig_calc_mode = xl.Calculation
    try:
        # switch Excel to manual calculation
        # and disable screen updating
        xl.Calculation = constants.xlManual
        xl.ScreenUpdating = False

        # run the minimization routine
        xl_obj_func = partial(obj_func, xl)
        minimize(xl_obj_func, X, method='nelder-mead')

    finally:
        # restore the original calculation
        # and screen updating mode
        xl.ScreenUpdating = True
        xl.Calculation = orig_calc_mode
        mbox = QMessageBox()
        mbox.setIcon(QMessageBox.Information)
        mbox.setText("Keep these optimization results?")
        mbox.setWindowTitle("Optimization Complete")
        mbox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        yes_no = mbox.exec_()
        if yes_no != QMessageBox.Ok:
            xl.Range('Inputs').Value = in_values

def obj_func(xl, arg):
    """Wraps a spreadsheet computation as a Python function."""
    # Copy argument values to input range
    xl.Range('Inputs').Value = [(float(x), ) for x in arg]

    # Calculate after changing the inputs
    xl.Calculate()

    # Return the value of the output cell
    result = xl.Range("Output").Value
    return result
