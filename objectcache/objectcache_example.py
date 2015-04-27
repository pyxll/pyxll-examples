"""
Object Cache Example
====================

Excel cells hold basic types (strings, numbers, booleans etc) but sometimes
it can be useful to have functions that take and return objects and to be
able to call those functions from Excel.

The functions in this module use the objectcache argument type and return
type defined in the objectcache module.

Both modules should be added to the pyxll.cfg file for the included
Excel example file.
"""
from pyxll import xl_func


class MyTestClass(object):
    """A basic class for testing the cached_object type"""

    def __init__(self, x):
        self.__x = x

    def __str__(self):
        return "%s(%s)" % (self.__class__.__name__, self.__x)


@xl_func("var: cached_object")
def cached_object_return_test(x):
    """returns an instance of MyTestClass"""
    return MyTestClass(x)


@xl_func("cached_object: string")
def cached_object_arg_test(x):
    """takes a MyTestClass instance and returns a string"""
    return str(x)


class MyDataGrid(object):
    """
    A second class for demonstrating cached_object types.
    This class is constructed with a grid of data and has
    some basic methods which are also exposed as worksheet
    functions.
    """

    def __init__(self, grid):
        self.__grid = grid

    def sum(self):
        """returns the sum of the numbers in the grid"""
        total = 0
        for row in self.__grid:
            total += sum(row)
        return total

    def __len__(self):
        total = 0
        for row in self.__grid:
            total += len(row)
        return total

    def __str__(self):
        return "%s(%d values)" % (self.__class__.__name__, len(self))


@xl_func("float[]: cached_object")
def make_datagrid(x):
    """returns a MyDataGrid object"""
    return MyDataGrid(x)


@xl_func("cached_object: int")
def datagrid_len(x):
    """returns the length of a MyDataGrid object"""
    return len(x)


@xl_func("cached_object: float")
def datagrid_sum(x):
    """returns the sum of a MyDataGrid object"""
    return x.sum()


@xl_func("cached_object: string")
def datagrid_str(x):
    """returns the string representation of a MyDataGrid object"""
    return str(x)


#
# Utility functions not to do with the objectcache exmaple, but
# used by the objectcache.xlsx spreadsheet.
#
@xl_func(": bool", volatile=True)
def win32com_is_installed():
    """returns True if win32com is installed"""
    try:
        import win32com
        return True
    except ImportError:
        return False


@xl_func("xl_cell: string", macro=True)
def get_formula(cell):
    """returns the formula of a cell"""
    return cell.formula
