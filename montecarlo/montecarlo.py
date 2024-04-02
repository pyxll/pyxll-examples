"""
Example showing how Monte Carlo style simulations can be achieved
in Excel using PyXLL.

This code accompanies the tutorial video
https://www.youtube.com/watch?v=Va9ih1DXPDs

- Random variables are created for multiple inputs
- Inputs cells are set to samples from those random variables
- An output cell is sampled each time after the inputs are set
- This is reapeated for a specified number of times and the results are collected

The mean and standard deviation are calculated from the list of results
and a histogram can be plotted using Seaborn.

Only one distribution is used in this example, the PERT distribution, but
other distributions can easily be added by writing further derived classes
of the `RandomVariable` class and adding more functions to create those from
Excel (see the function `mc_pert`).
"""
from pyxll import xl_func, xl_macro, XLCell, plot, xl_app, xlcAlert
from contextlib import contextmanager, ExitStack
from abc import ABCMeta, abstractmethod
from win32com.client import constants
from pert import PERT
from functools import wraps
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from pywin.mfc import dialog
import win32con, win32ui
import threading
import queue


@contextmanager
def disable_auto_calc():
    """Context manager to disable automatic calculation and screen updates."""
    xl = xl_app()
    calc_mode = xl.Calculation
    try:
        xl.ScreenUpdating = False
        xl.Calculation = constants.xlManual
        yield
    finally:
        xl.Calculation = calc_mode
        xl.ScreenUpdating = True


def alert_on_error(func):
    """Decorator to display an alert if an exception is raised.
    Uses pyxll.xlcAlert so can only be used for macro functions.
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            func(*args, **kwargs)
        except Exception as e:
            xlcAlert(f"An error occurred:\n\n"
                     f"{type(e).__name__}: {e}\n\n"
                     f"Check the log for details")
            raise
    return wrapper


# Adapted from https://www.pyxll.com/blog/a-super-simple-excel-progress-bar/
@contextmanager
def progress_bar(title="Working..."):
    """Context manager for a progress indicator bar.

    Yields a function that takes the current progress as a number between 0 and 1,
    and returns False if the cancel button has been pressed.
    """
    style = (
        win32con.DS_MODALFRAME |
        win32con.WS_POPUP |
        win32con.WS_VISIBLE |
        win32con.WS_CAPTION |
        win32con.WS_SYSMENU |
        win32con.DS_SETFONT
    )

    button_style = (
        win32con.WS_CHILD |
        win32con.WS_VISIBLE |
        win32con.WS_TABSTOP |
        win32con.BS_PUSHBUTTON
    )

    w = 215
    h = 36

    template = [
        # Dialog
        [
            title,
            (0, 0, w, h),
            style,
            None,
            (8, "MS Sans Serif")
        ],
        # Cancel button
        [
            0x80,  # Button
            "Cancel",
            win32con.IDCANCEL,
            (w - 60, h - 18, 50, 14),
            button_style
        ]
    ]

    class ProgressDialog(dialog.Dialog):
        def __init__(self, template):
            super().__init__(template)
            self.__closed = False

        def OnInitDialog(self):
            rc = dialog.Dialog.OnInitDialog(self)
            self.pbar = win32ui.CreateProgressCtrl()
            self.pbar.CreateWindow(
                win32con.WS_CHILD | win32con.WS_VISIBLE,
                (10, 10, 460, 24),
                self, 1001
            )
            self.pbar.SetRange(1, 100)
            return rc

        def OnCancel(self):
            super().OnCancel()
            self.__closed = True

        def set_progress(self, progress):
            if not self.__closed:
                self.pbar.SetPos(progress)
            return not self.__closed
    
        def close(self):
            if not self.__closed:
                self.PostMessage(win32con.WM_CLOSE, 0, 0)

    def show_progress_dialog(q):
        # Create the progress dialog window
        dlg = ProgressDialog(template)

        # Pass the new dialog back to the main thread
        q.put(dlg)

        # And display the dialog
        dlg.DoModal()

    # Create and start a background thread to display the progress indicator
    q = queue.Queue()
    thread = threading.Thread(target=show_progress_dialog, args=(q,))
    thread.daemon = True
    thread.start()

    # Wait for the dialog object created by the background thread
    dlg = q.get(timeout=5)

    try:
        # Yield a function to set the progress, between 0 and 1
        yield lambda x: dlg.set_progress(int(max(min(x, 1), 0) * 100))
    finally:
        # Close the dialog and wait for the thread to end
        dlg.close()
        thread.join(timeout=5)


class RandomVariable(metaclass=ABCMeta):
    """Base class for random variables."""

    def __init__(self, name: str, target: XLCell):
        self.name = name
        self.target = target

    @abstractmethod
    def samples(self, n, seed=None):
        """Return n random samples."""
        pass

    def __enter__(self):
        self.__original_value = self.target.value

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.target.value = self.__original_value

    def simulate(self, value):
        self.target.value = value


class PertRandomVariable(RandomVariable):
    """Random variable using the PERT distribution."""

    def __init__(self,
                 name: str,
                 input: XLCell,
                 min_value: float,
                 ml_value: float,
                 max_vaue: float,
                 lamb: float = 4.0):
        super().__init__(name, input)
        self.__dist = PERT(min_value, ml_value, max_vaue, lamb)

    def samples(self, n, seed=None):
        return self.__dist.rvs(size=n, random_state=seed)


class Simulation:
    """Simulation class for managing setting inputs to random variables
    and collecting the calculated output.
    """

    def __init__(self,
                 name: str,
                 n: int,
                 output: XLCell,
                 inputs: list[RandomVariable]):
        self.name = name
        self.n = n
        self.output = output
        self.inputs = inputs

    def run(self, seed=None, set_progress=None):
        output_range = self.output.to_range()

        results = []
        with ExitStack() as stack:
            # Restore the input values when finished
            for input in self.inputs:
                stack.enter_context(input)

            # Prepare random samples for the inputs
            samples = [input.samples(self.n, seed=seed) for input in self.inputs]

            # Calculate the output value for each set of inputs
            for i, values in enumerate(zip(*samples)):
                for input, value in zip(self.inputs, values):
                    input.simulate(value)

                output_range.Calculate()
                results.append(self.output.value)

                if set_progress is not None and 0 == i % 100:
                    if not set_progress(i / self.n):
                        raise RuntimeError("Calculation cancelled")

        return results


@xl_func("string, xl_cell, float, float, float, float: object")
def mc_pert(name, input, min_value, ml_value, max_vaue, lamb=4.0):
    """Returns a RandomVariable object instance using the PERT distribution."""
    return PertRandomVariable(name, input, min_value, ml_value, max_vaue, lamb)


@xl_func("object rv, int n, int seed, str style, str context: str")
def mc_plot_rv(rv, n=10000, seed=None, style="whitegrid", context="paper"):
    """Plot the distribution of a random variable."""
    with sns.axes_style(style=style), \
            sns.plotting_context(context=context):
        samples = rv.samples(n=n, seed=seed)
        fig, ax = plt.subplots()
        sns.kdeplot(samples, ax=ax)
        ax.set(title=rv.name)
        plot(fig)

    return f"[{rv.name}]"


@xl_func("str name, int n, xl_cell output, object *inputs")
def mc_simulation(name, n, output, *inputs):
    """Returns a Simulation object for use with the 'mc_run_simulation' macro."""
    return Simulation(name, n, output, inputs)


@xl_macro
@alert_on_error
def mc_run_simulation():
    """Run a Simulation located at the cell specified in the calling button's
    alternative text field.
    """
    # Get the Simulation object from the cell
    xl = xl_app()
    caller = xl.Caller
    button = xl.ActiveSheet.Shapes[caller]
    address = button.AlternativeText

    cell = XLCell.from_range("C14")

    cell = XLCell.from_range(address)
    simulation = cell.options(type="object").value
    if not isinstance(simulation, Simulation):
        raise RuntimeError("Cell value is not a Simulation object")

    # Run the simulation with automatic calculations disabled and
    # display a progress indicator.
    with disable_auto_calc(), \
            progress_bar() as set_progress:
        results = simulation.run(set_progress=set_progress)

    # Write the results as an object to the cell below
    cell.offset(1, 0).options(type="object").value = results


@xl_func("object simulation, object results, str style, str context: str")
def mc_plot_sim(simulation, results, style="whitegrid", context="paper"):
    """Plot the results of a simulation run."""
    if not results:
        return "# No results"

    with sns.axes_style(style=style), \
            sns.plotting_context(context=context):
        fig, ax = plt.subplots()
        sns.histplot(results, ax=ax)
        ax.set(title=simulation.name)
        plot(fig)

    return f"[{simulation.name}]"


@xl_func("object results: float")
def mc_mean(results):
    """Return the mean of a list of results."""
    if not results:
        return "# No results"
    return np.mean(results)


@xl_func("object results: float")
def mc_stddev(results):
    """Return the standard deviation of a list of results."""
    if not results:
        return "# No results"
    return np.std(results)
