"""
Profiling tools for use with PyXLL.

This code accompanies the blog post
https://www.pyxll.com/blog/how-to-profile-python-code-in-excel
"""
from pyxll import xl_menu, xl_app, xlcAlert
from win32com.client import constants
from functools import wraps
import win32clipboard
import line_profiler
import cProfile
import pstats
import math
import time

try:
    from io import StringIO
except ImportError:
    from StringIO import StringIO


_active_cprofiler = None
_active_line_profiler = None


@xl_menu("Time to Calculate", menu="Profiling Tools")
def time_calculation():
    """Recalcules the selected range and times how long it takes"""
    xl = xl_app()

    # switch Excel to manual calculation
    orig_calc_mode = xl.Calculation
    try:
        xl.Calculation = constants.xlManual

        # get the current selection and its formula
        selection = xl.Selection

        # run the calculation a few times
        timings = []
        for i in range(100):
            # Start the timer and set the selection formula to itself.
            # This is a reliable way to force Excel to recalculate the range.
            start_time = time.clock()
            selection.Calculate()
            end_time = time.clock()
            duration = end_time - start_time
            timings.append(duration)

        # calculate the mean and stddev
        mean = math.fsum(timings) / len(timings)
        stddev = (math.fsum([(x - mean) ** 2 for x in timings]) / len(timings)) ** 0.5
        best = min(timings)
        worst = max(timings)

        # copy the results to the clipboard
        data = [
            ["mean", mean],
            ["stddev", stddev],
            ["best", best],
            ["worst", worst]
        ]
        text = "\n".join(["\t".join(map(str, x)) for x in data])
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()

        # report the results
        xlcAlert(("%0.2f ms \xb1 %d \xb5s\n"
                  "Best: %0.2f ms\n"
                  "Worst: %0.2f ms\n"
                  "(Copied to clipboard)") % (mean * 1000, stddev * 1000000, best * 1000, worst * 1000))
    finally:
        # restore the original calculation mode
        xl.Calculation = orig_calc_mode


@xl_menu("Start", menu="Profiling Tools", sub_menu="cProfile")
def start_profiling():
    """Start the cProfile profiler"""
    global _active_cprofiler
    if _active_cprofiler is not None:
        _active_cprofiler.disable()
    _active_cprofiler = cProfile.Profile()

    xlcAlert("cProfiler Active\n"
             "Recalcuate the workbook and then stop the profiler\n"
             "to see the results.")

    _active_cprofiler.enable()


@xl_menu("Stop", menu="Profiling Tools", sub_menu="cProfile")
def stop_profiling():
    """Stop the cProfile profiler and print the results"""
    global _active_cprofiler
    if not _active_cprofiler:
        xlcAlert("No active profiler")
        return

    _active_cprofiler.disable()

    # print the profiler stats
    stream = StringIO()
    stats = pstats.Stats(_active_cprofiler, stream=stream).sort_stats("cumulative")
    stats.print_stats()

    # print the results to the log
    print(stream.getvalue())

    # and copy to the clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(stream.getvalue())
    win32clipboard.CloseClipboard()

    _active_cprofiler = None

    xlcAlert("cProfiler Stopped\n"
             "Results have been written to the log and clipboard.")


# Current active line profiler
_active_line_profiler = None


def enable_line_profiler(func):
    """Decorator to switch on line profiling for a function.

    If using line_profiler from the command line, use the built-in @profile
    decorator instead of this one.
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        nonlocal func
        if _active_line_profiler:
            func = _active_line_profiler(func)
        return func(*args, **kwargs)
    return wrapper


@xl_menu("Start", menu="Profiling Tools", sub_menu="Line Profiler")
def start_line_profiler():
    """Start the line profiler"""
    global _active_line_profiler
    _active_line_profiler = line_profiler.LineProfiler()

    xlcAlert("Line Profiler Active\n"
             "Run the function you are interested in and then stop the profiler.\n"
             "Ensure you have decoratored the function with @enable_line_profiler.")


@xl_menu("Stop", menu="Profiling Tools", sub_menu="Line Profiler")
def stop_line_profiler():
    """Stops the line profiler and prints the results"""
    global _active_line_profiler
    if not _active_line_profiler:
        return

    stream = StringIO()
    _active_line_profiler.print_stats(stream=stream)
    _active_line_profiler = None

    # print the results to the log
    print(stream.getvalue())

    # and copy to the clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(stream.getvalue())
    win32clipboard.CloseClipboard()

    xlcAlert("Line Profiler Stopped\n"
             "Results have been written to the log and clipboard.")


