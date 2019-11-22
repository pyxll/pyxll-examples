"""
Optimisation example using scipy.optimize.minimize.

This example shows how to find a minimum for a function
of two variables, though the principles are very easily
extended to N.

This version uses named ranges to access the input and
output cells of the objective function.

This code accompanies the blog post
  https://www.pyxll.com/blog/a-better-goal-seek/
"""
from functools import partial
from pyxll import xl_macro, xl_app
from win32com.client import constants
import numpy as np
from scipy.optimize import minimize

@xl_macro(shortcut="Ctrl+Alt+R")
def optimize3():
    xl = xl_app()
    # Get the initial values of the input cells
    in_values = xl.Range('Inputs').Value
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


def obj_func(xl, arg):
    """Wraps a spreadsheet computation as a Python function."""
    # Copy argument values to input range
    xl.Range('Inputs').Value = [(float(x), ) for x in arg]

    # Calculate after changing the inputs
    xl.Calculate()

    # Return the value of the output cell
    result = xl.Range("Output").Value
    return result
