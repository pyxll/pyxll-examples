"""
Optimisation example using scipy.optimize.minimize.

Extended to query the user for input and output cells.

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
from op_dialog import OpDialog


def get_qt_app():
    """Returns a QApplication instance.

    MUST be called before showing any dialogs.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([sys.executable])
    return app


def get_range(s):
    xl = xl_app()
    try:
        return xl.Range(s)
    except Exception as e:
        raise ValueError("Range specification not acceptable")


@xl_menu("Optimize5")
def optimize5():
    """
    Trigger optimization of a spreadsheet model that
    takes the named range "Inputs" as inputs and
    produces output in the named range "Output".
    """
    xl = xl_app()
    qt_app = get_qt_app()  # pragma noqc
    # Get the initial values of the input cells
    msgBox = OpDialog()
    result = msgBox.exec_()
    if not result:  # user cancelled
        return

    in_range = get_range(msgBox.in_range.text())
    out_cell = get_range(msgBox.out_cell.text())
    in_values = list(in_range.Value)
    X = np.array([x[0] for x in in_values])

    orig_calc_mode = xl.Calculation
    try:
        # switch Excel to manual calculation
        # and disable screen updating
        xl.Calculation = constants.xlManual
        xl.ScreenUpdating = False

        # run the minimization routine
        xl_obj_func = partial(obj_func, xl, in_range, out_cell)
        print(f"X = {X}")
        result = minimize(xl_obj_func, X, method="nelder-mead")
        in_range.Value = [(float(x),) for x in result.x]
        xl.ScreenUpdating = True
        mbox = QMessageBox()
        mbox.setIcon(QMessageBox.Information)
        mbox.setText("Optimization results shown below." "\nMake changes permanent?")
        mbox.setWindowTitle("Optimization Complete")
        mbox.setInformativeText(
            "\n".join(
                [
                    "Successful:       %s" % result.success,
                    result.message,
                    "After %d iterations" % result.nit,
                ]
            )
        )
        mbox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        yes_no = mbox.exec_()
        if yes_no != QMessageBox.Ok:
            in_range.Value = in_values
        else:
            in_range.Value = [(float(x),) for x in result.x]

    finally:
        # restore the original calculation
        # and screen updating mode
        xl.ScreenUpdating = True
        xl.Calculation = orig_calc_mode


def obj_func(xl, in_range, out_cell, arg):
    """Wraps a spreadsheet computation as a Python function."""
    # Copy argument values to input range
    in_range.Value = [(float(x),) for x in arg]

    # Calculate after changing the inputs
    xl.Calculate()

    # Return the value of the output cell
    result = float(out_cell.Value)
    return result
