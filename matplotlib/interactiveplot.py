"""
Example code showing how to draw an interactive matplotlib figure
in Excel.

While the figure is displayed Excel is still useable in the background
and the chart may be updated with new data by calling the same
function again.
"""
from pyxll import xl_func

from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt4agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
import matplotlib as mpl
import matplotlib.pyplot as plt

from PySide import QtCore, QtGui
import timer

# dict to keep track of any chart windows
_plot_windows = {}


@xl_func("string figname, numpy_column<float> xs, numpy_array<float> ys, var style: string")
def mpl_line_plot(figname, xs, ys, style=None):
    """
    Do a matplotlib line plot in an interactive window.

    :param figname: name to use for this plot's window
    :param xs: list of x values as a column
    :param ys: 2d array of y values, arranged as columns
    :param style: plot style, eg 'fivethirtyeight' (optional)
    """
    # Get the Qt app.
    # Note: no need to 'exec' this as it will be polled in the main windows loop.
    app = get_qt_app()

    # if using a style get the current settings and restore them after plotting
    if style is not None:
        initial_settings = mpl.rcParams.copy()
        plt.style.use(style)

    try:
        # generate the plot
        fig = Figure(figsize=(600, 600), dpi=72, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
        ax = fig.add_subplot(111)
        ax.plot(xs, ys)
    finally:
        # restore any settings after plotting (this is only necessary if styling)
        if style:
            mpl.rcParams.update(initial_settings)

    # generate the canvas to display the plot
    canvas = FigureCanvas(fig)
 
    # Get or create the Qt windows to show the chart in.
    if figname in _plot_windows:
        # get from the global dict and clear any previous widgets
        window = _plot_windows[figname]
        layout = window.layout()
        if layout:
            for i in reversed(range(layout.count())):
                layout.itemAt(i).widget().setParent(None)
    else:
        # create a new window for this plot and store it for next time
        window = QtGui.QWidget()
        window.resize(800, 600)
        window.setWindowTitle(figname)
        _plot_windows[figname] = window

    # create the navigation toolbar
    toolbar = NavigationToolbar(canvas, window)

    # add the canvas and toolbar to the window
    layout = window.layout() or QtGui.QVBoxLayout()
    layout.addWidget(canvas)
    layout.addWidget(toolbar)
    window.setLayout(layout)

    window.show()
    return "[Plotted '%s']" % figname


#
# Taken from the ui/qt.py example
#
def get_qt_app():
    """
    returns the global QtGui.QApplication instance and starts
    the event loop if necessary.
    """
    app = QtCore.QCoreApplication.instance()
    if app is None:
        # create a new application
        app = QtGui.QApplication([])

        # use timer to process events periodically
        processing_events = {}
        def qt_timer_callback(timer_id, time):
            if timer_id in processing_events:
                return
            processing_events[timer_id] = True
            try:
                app = QtCore.QCoreApplication.instance()
                if app is not None:
                    app.processEvents(QtCore.QEventLoop.AllEvents, 300)
            finally:
                del processing_events[timer_id]

        timer.set_timer(100, qt_timer_callback)

    return app
