"""
Functions used to demonstrate performance profiling.

This code accompanies the blog post
https://www.pyxll.com/blog/how-to-profile-python-code-in-excel
"""
from pyxll import xl_func
import numba


@xl_func
def average_tol_simple(values, tolerance):
    """Compute the mean of values where value is > tolerance
    
    :param values: Range of values to compute average over
    :param tolerance: Only values greater than tolerance will be considered
    """
    total = 0.0
    count = 0
    for row in values:
        for value in row:
            if not isinstance(value, (float, int)):
                continue
            if value <= tolerance:
                continue
            total += value
            count += 1
    return total / count


@xl_func
def average_tol_try_except(values, tolerance):
    """Compute the mean of values where value is > tolerance

    :param values: Range of values to compute average over
    :param tolerance: Only values greater than tolerance will be considered
    """
    total = 0.0
    count = 0
    for row in values:
        for value in row:
            try:
                if value <= tolerance:
                    continue
                total += value
                count += 1
            except TypeError:
                continue
    return total / count


@xl_func
@numba.jit(locals={"value": numba.double, "total": numba.double, "count": numba.int32})
def average_tol_numba(values, tolerance):
    """Compute the mean of values where value is > tolerance

    :param values: Range of values to compute average over
    :param tolerance: Only values greater than tolerance will be considered
    """
    total = 0.0
    count = 0
    for row in values:
        for value in row:
            if value <= tolerance:
                continue
            total += value
            count += 1
    return total / count


@xl_func("numpy_array<float>, float: float")
def average_tol_numpy(values, tolerance):
    """Compute the mean of values where value is > tolerance

    :param values: Range of values to compute average over
    :param tolerance: Only values greater than tolerance will be considered
    """
    return values[values > tolerance].mean()
