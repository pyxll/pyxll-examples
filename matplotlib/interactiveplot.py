"""
Example code showing how to draw an interactive matplotlib figure
in Excel.

While the figure is displayed Excel is still useable in the background
and the chart may be updated with new data by calling the same
function again.
"""
from pyxll import xl_func
from pandas.stats.moments import ewma

# matplotlib imports
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt4agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

# Qt imports
from PySide import QtCore, QtGui
import timer  # for polling the Qt application

# dict to keep track of any chart windows
_plot_windows = {}

@xl_func("string figname, numpy_column<float> xs, numpy_column<float> ys, int span: string")
def mpl_plot_ewma(figname, xs, ys, span):
    """
    Show a matplotlib line plot of xs vs ys and ewma(ys, span) in an interactive window.

    :param figname: name to use for this plot's window
    :param xs: list of x values as a column
    :param ys: list of y values as a column
    :param span: ewma span
    """
    # Get the Qt app.
    # Note: no need to 'exec' this as it will be polled in the main windows loop.
    app = get_qt_app()

    # create the figure and axes for the plot
    fig = Figure(figsize=(600, 600), dpi=72, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
    ax = fig.add_subplot(111)

    # calculate the moving average
    ewma_ys = ewma(ys, span=span)

    # plot the data
    ax.plot(xs, ys, alpha=0.4, label="Raw")
    ax.plot(xs, ewma_ys, label="EWMA")
    ax.legend()

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
