"""
Optimisation example using scipy.optimize.minimize.

This example shows how to find a minimum for a function
of two variables, though the principles are very easily
extended to N.

This version uses ranges to access the input values
rather than loading each cell individually.

This code accompanies the blog post
  https://www.pyxll.com/blog/a-better-goal-seek/
"""
from pyxll import xl_macro, xl_app
from win32com.client import constants
import numpy as np
from scipy.optimize import minimize

@xl_macro(shortcut="Ctrl+Alt+Q")
def optimize2():
    xl = xl_app()
    # Get the initial values of the input cells
    in_values = xl.Range('C11:C12').Value
    X = np.array([x[0] for x in in_values])

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
    xl.Range('C11:C12').Value = [(float(x), ) for x in arg]

    # Calculate after changing the inputs
    xl.Calculate()

    # Return the value of the output cell
    result = xl.Range("E11").Value
    return result

