"""
Example code showing how to draw an interactive matplotlib figure
in Excel.

This uses a right click context menu that can be used on any
cell returning a matplotlib Figure object.

The Figure is displayed in a Tk window using PyXLL's Custom
Task Pane feature.

The menu item is configured in the 'ribbon.xml' found in the same
folder as this file.

This module also contains a worksheet function 'show_matplotlib_ctp'
which can be called from a worksheet to display a matplotlib Figure
in a CTP. This second function shows how the widget contained in the
CTP can be updated as inputs change and the function is re-run.
"""
from pyxll import xl_func, create_ctp, xl_app, xlfCaller, XLCell, CTPDockPositionFloating
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
from weakref import WeakValueDictionary
from PySide6.QtWidgets import QWidget, QVBoxLayout, QApplication


def show_selected_matplotlib_ctp(control):
    """Context menu action callback.

    Gets a matplotlib Figure object from the current Excel selection
    and displays in in a custom task pane.
    """
    # Get the Excel application object
    xl = xl_app()

    # Get the current and check if it as a matplotlib Figure
    cell = XLCell.from_range(xl.Selection)
    fig = cell.options(type="object").value

    if not isinstance(fig, Figure):
        raise ValueError("Expected a matplotlib Figure object")

    # Before we can create a Qt widget the Qt App must have been initialized.
    # Make sure we keep a reference to this until create_ctp is called.
    app = QApplication.instance()
    if app is None:
        app = QApplication([])

    # Create the widget and a layout for it
    widget = QWidget()
    layout = QVBoxLayout(widget)
    widget.setLayout(layout)

    # Add the matplotlib plot to the window
    canvas = FigureCanvas(fig)
    widget.layout().addWidget(canvas)

    # And add a toolbar
    toolbar = NavigationToolbar(canvas)
    widget.layout().addWidget(toolbar)

    # Show as a custom task pane using PyXLL.create_ctp
    create_ctp(widget, width=800, height=800, position=CTPDockPositionFloating)


# Dictionary of calling cell addresses to Qt widgets
_mpl_ctp_cache = WeakValueDictionary()


@xl_func("object fig, bool enabled: var", macro=True)  # macro=True is needed for xlfCaller
def show_matplotlib_ctp(fig, enabled=True):
    """Display a matplotlib Figure in a Custom Task Pane.

    This worksheet function takes a cell reference rather than
    an object directly as it keeps track of the custom task pane
    and updates it with the new figure if called again for the same
    cell.
    """
    if not enabled:
        return fig

    if not isinstance(fig, Figure):
        raise ValueError("Expected a matplotlib Figure object")

    # Get the calling cell to check if there is already a visible CTP for this cell
    cell = xlfCaller()

    # Get the widget from the cache if it exists already
    widget = _mpl_ctp_cache.get(cell.address, None)
    show_ctp = True if widget is None else False
    if widget is None:
        # Before we can create a Qt widget the Qt App must have been initialized.
        # Make sure we keep a reference to this until create_ctp is called.
        app = QApplication.instance()
        if app is None:
            app = QApplication([])

        # Create the widget and a layout for it
        widget = QWidget()
        layout = QVBoxLayout(widget)
        widget.setLayout(layout)

    # Close any old widgets and remove them from the layout
    layout = widget.layout()
    while layout.count() > 0:
        child = layout.itemAt(0)
        child.widget().close()
        layout.removeItem(child)

    # Add the matplotlib plot to the window
    canvas = FigureCanvas(fig)
    widget.layout().addWidget(canvas)

    # And add a toolbar
    toolbar = NavigationToolbar(canvas)
    widget.layout().addWidget(toolbar)

    # Create and show the CTP if necessary
    if show_ctp:
        create_ctp(widget, width=800, height=800, position=CTPDockPositionFloating)

    # We use a WeakValueDict so the item stays in this dict so long as the widget is alive.
    # Once the CTP is closed and the widget is destroyed then the item in the cache is
    # cleared automatically.
    _mpl_ctp_cache[cell.address] = widget

    return fig
