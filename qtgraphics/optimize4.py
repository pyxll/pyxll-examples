"""
Optimisation example using scipy.optimize.minimize.

Extended to ad.

This code accompanies the blog post
  https://www.pyxll.com/blog/extending-the-excel-user-interface/
"""
import sys

from functools import partial
from pyxll import xl_app, xl_menu
from win32com.client import constants
import numpy as np
from scipy.optimize import minimize
from PyQt5.QtWidgets import QApplication, QMessageBox


def get_qt_app():
    """Returns a QApplication instance.

    MUST be called before showing any dialogs.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([sys.executable])
    return app


@xl_menu("Optimize")
def optimize4():
    """
    Trigger optimization of a spreadsheet model that
    takes the named range "Inputs" as inputs and
    produces output in the named range "Output".
    """
    qt_app = get_qt_app()  # pragma noqc
    app = QApplication.instance()
    if app is None:
        app = QApplication([sys.executable])
    return app
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
        result = minimize(xl_obj_func, X, method='nelder-mead')
        mbox = QMessageBox()
        mbox.setIcon(QMessageBox.Information)
        mbox.setText("Optimization results shown below."
                     "\nMake changes permanent?")
        mbox.setWindowTitle("Optimization Complete")
        mbox.setInformativeText("\n".join([
            "Successful:       %s" % result.success,
            result.message,
            "After %d iterations" % result.nit
            ]))
        mbox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        yes_no = mbox.exec_()
        if yes_no != QMessageBox.Ok:
            xl.Range('Inputs').Value = in_values
        else:
            xl.Range('Inputs').Value = [(float(x), ) for x in result.x]

    finally:
        # restore the original calculation
        # and screen updating mode
        xl.ScreenUpdating = True
        xl.Calculation = orig_calc_mode


def obj_func(xl, arg):
    """Wraps a spreadsheet computation as a Python function."""
    # Copy argument values to input range
    xl.Range('Inputs').Value = [(float(x), ) for x in arg]

    # Calculate after changing the inputs
    xl.Calculate()

    # Return the value of the output cell
    result = xl.Range("Output").Value
    return result
