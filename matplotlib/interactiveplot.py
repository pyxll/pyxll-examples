"""
Example code showing how to draw an interactive matplotlib figure
in Excel.

This uses a right click context menu that can be used on any
cell returning a matplotlib Figure object.

The Figure is displayed in a Tk window using PyXLL's Custom
Task Pane feature.

The menu item is configured in the 'ribbon.xml' found in the same
folder as this file.
"""
from pyxll import create_ctp, xl_app, XLCell, CTPDockPositionFloating
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import tkinter as tk


def show_matplotlib_ctp(control):
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

    # Create the top level Tk window and give it a title
    root = tk.Toplevel()
    root.title("Matplotlib Plot")

    # Add the matplotlib plot to the window
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()

    # Add a toolbar to the layout
    toolbar = NavigationToolbar2Tk(canvas, root, pack_toolbar=False)
    toolbar.update()
    toolbar.pack(side=tk.BOTTOM, fill=tk.X)
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    # Show as a custom task pane using PyXLL.create_ctp
    create_ctp(root, width=800, height=800, position=CTPDockPositionFloating)
