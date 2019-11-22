"""
Optimisation example using scipy.optimize.minimize.

This example shows how to find a minimum for a function
of two variables, though the principles are very easily
extended to N.

This code accompanies the blog post
  https://www.pyxll.com/blog/a-better-goal-seek/
"""
from pyxll import xl_macro, xl_app
from scipy.optimize import minimize
from win32com.client import constants
import numpy as np

@xl_macro(shortcut="Ctrl+Alt+P")
def optimise1():
    xl = xl_app()
    # Get the initial values of the input cells
    x = xl.Range("C11").Value
    y = xl.Range("C12").Value
    X = np.array([x, y])

    orig_calc_mode = xl.Calculation
    try:
        # switch Excel to manual calculation
        # and disable screen updating
        xl.Calculation = constants.xlManual
        xl.ScreenUpdating = False

        # run the minimization routine
        minimize(obj_func, X, method='nelder-mead')

    finally:
        # restore the original calculation
        # and screen updating mode
        xl.ScreenUpdating = True
        xl.Calculation = orig_calc_mode


def obj_func(arg):
    """Wraps a spreadsheet computation as a Python function."""
    xl = xl_app()
    # Copy argument values to input range
    xl.Range('C11').Value = float(arg[0])
    xl.Range('C12').Value = float(arg[1])

    # Calculate after changing the inputs
    xl.Calculate()

    # Return the value of the output cell
    result = xl.Range("E11").Value
    return result

