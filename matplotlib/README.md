# Python Plotting in Excel with Matplotlib

Example showing how to display a matplotlib figure in Excel.

The 'embeddedplot.py' example demonstrates how to display a matplotlib figure
directly in the Excel sheet using PyXLL's ``plot`` function.

More details can be found in the PyXLL docs https://www.pyxll.com/docs/userguide/plotting/index.html.

The 'interactiveplot.py' example shows how it's possible to add a menu item to Excel's
right click context menu. The added menu item, when called, fetches the matplotlib
Figure object from the currently selected cell (e.g. returned from the ``mpl_plot_ewma``
function from 'embeddedplot.py') and displays it in an Excel window.

This uses PyXLL's 'Custom Task Pane' feature which is documented here https://www.pyxll.com/docs/userguide/ctps/index.html

Also see https://www.pyxll.com/docs/userguide/contextmenus.html for information about
adding context menu items.
