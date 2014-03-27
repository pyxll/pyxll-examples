"""
Example Excel macro bound to a shortcut key.

The macro 'resize_array_formula' looks at the currently selected cell
and if it's an array formula re-calculates that cell and resizes the
output range so that it matches the dimensions of the result of the array
formula.

This can be useful when dealing with array formulas where the dimensions
of the result can vary and it's tedious to have to keep resizing the
Excel formula manually.

The function is registered as a menu item as well as a macro and can
be run from the menu or via the keyboard shortcut Ctrl+Shift+R.
"""
from pyxll import xl_menu, xl_macro, get_active_object
from shortcuts import xl_shortcut
from pywintypes import com_error
from shortcuts import xl_shortcut
from win32com.client import Dispatch
import win32api
import win32con
import logging

_log = logging.getLogger(__name__)


@xl_shortcut("Ctrl+Shift+R")
@xl_menu("Resize Array Formula (Ctrl+Shift+R)")
@xl_macro("")
def resize_array_formula():
    """
    Recalculates and resizes a range to show all the results of a formula.
    """
    xl_window = get_active_object()
    xl = Dispatch(xl_window).Application

    selection = xl.Selection
    formula = selection.FormulaArray

    if not (formula and (formula.startswith("=") or formula.startswith("+"))):
        # nothing to do
        return 

    # get the range of the entire formula
    current_range = _expand_range(xl, selection)

    # evaluate the formula to get the dimensions of the result
    # (this is an optimization to avoid converting the range into a python list)
    # result = xl.Evaluate(formula)
    result = xl._oleobj_.InvokeTypes(1, 0, 1, (12, 0), ((12, 1),), formula)
    
    width, height = 0, len(result)
    if height > 0 and isinstance(result[0], (list, tuple)):
        width = len(result[0])
    width, height = max(width, 1), max(height, 1)

    new_range = xl.Range(current_range.Offset(1, 1),
                         current_range.Offset(height, width))

    # check if we're overwriting any existing data
    if new_range.Count > current_range.Count:
        current_non_blank = current_range.Count - xl.WorksheetFunction.CountBlank(current_range)
        new_non_blank = new_range.Count - xl.WorksheetFunction.CountBlank(new_range)
        if current_non_blank != new_non_blank:
            new_range.Select()
            result = win32api.MessageBox(None,
                                         "Content will be overwritten in %s" % new_range.Address,
                                         "Warning",
                                         win32con.MB_OKCANCEL | win32con.MB_ICONWARNING)
            if result == win32con.IDCANCEL:
                current_range.FormulaArray = formula
                return

    # clear the old range
    current_range.ClearContents()

    # set the formula on the new range
    try:
        new_range.FormulaArray = formula
    except Exception:
        result = win32api.MessageBox(None,
                                     "Error resizing range",
                                     "Error",
                                     win32con.MB_OKCANCEL | win32con.MB_ICONERROR)
        _log.error("Error resizing range", exc_info=True)
        current_range.FormulaArray = formula


def _expand_range(xl, selection):
    """
    Expand a range to include the whole formula range, if the initial
    range is part of an array formula.
    """
    formula = selection.FormulaArray
    if not (formula and (formula.startswith("=") or formula.startswith("+"))):
        # nothing to do
        return  selection
    
    # Range.Offset if weird as a value of 1 results in no offset
    idx_offset = 1
    # find if there are any cells above the selection with the same formula
    top_left = selection
    try:
        if top_left.Offset(idx_offset-1).FormulaArray == formula:
            # select up (this could go outside the formula range)
            top_left = top_left.End(constants.xlUp)

            # move down until we find the formula
            while top_left.FormulaArray != formula:
                top_left = top_left.Offset(idx_offset+1)
    except com_error:
        pass

    try:
        if top_left.Offset(idx_offset+0, idx_offset-1).FormulaArray == formula:
            # select left (this could go outside the formula range)
            top_left = top_left.End(constants.xlToLeft)

            # move right until we find the formula
            while top_left.FormulaArray != formula:
                top_left = top_left.Offset(idx_offset+0, idx_offset+1)
    except com_error:
        pass

    bottom_right = selection
    try:
        if bottom_right.Offset(idx_offset+1).FormulaArray == formula:
            # select down (this could go outside the formula range)
            bottom_right = bottom_right.End(constants.xlDown)

            # move up until we find the formula
            while bottom_right.FormulaArray != formula:
                bottom_right = bottom_right.Offset(idx_offset-1)
    except com_error:
        pass

    try:
        if bottom_right.Offset(idx_offset+0, idx_offset+1).FormulaArray == formula:
            # select right (this could go outside the formula range)
            bottom_right = bottom_right.End(constants.xlToRight)

            # move left until we find the formula
            while bottom_right.FormulaArray != formula:
                bottom_right = bottom_right.Offset(idx_offset+0, idx_offset-1)
    except com_error:
        pass

    return xl.Range(top_left, bottom_right)
