"""
Example code showing how to draw a static matplotlib figure
in Excel.

Matplotlib is used to plot a chart to an image, which is then
displayed as an object in Excel.
"""
from pyxll import xl_func, xlfCaller, async_call
from pandas.stats.moments import ewma
import os

# matplotlib imports
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.figure import Figure

# For interacting with Excel from Python
from pyxll import get_active_object
import win32com.client


@xl_func("string figname, "
         "numpy_column<float> xs, "
         "numpy_column<float> ys, "
         "int span: string",
         macro=True)
def mpl_plot_ewma_embedded(figname, xs, ys, span):
    # create the figure and axes for the plot
    fig = Figure(figsize=(8, 6), dpi=150, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
    ax = fig.add_subplot(111)

    # calculate the moving average
    ewma_ys = ewma(ys, span=span)

    # plot the data
    ax.plot(xs, ys, alpha=0.4, label="Raw")
    ax.plot(xs, ewma_ys, label="EWMA")
    ax.legend()

    # write the figure to a temporary image file
    filename = os.path.join(os.environ["TEMP"], "xlplot_%s.png" % figname)
    canvas = FigureCanvas(fig)
    canvas.draw()
    canvas.print_png(filename)

    # Show the figure in Excel as a Picture object on the same sheet
    # the function is being called from.
    xl = xl_app()
    caller = xlfCaller()
    sheet = xl.Range(caller.address).Worksheet

    # insert the picture
    picture = sheet.Pictures().Insert(filename)

    # if a picture with the same figname already exists then resize
    # the new picture and delete the old one
    for old_picture in sheet.Pictures():
        if old_picture.Name == figname:
            picture.Height = old_picture.Height
            picture.Width = old_picture.Width
            picture.Top = old_picture.Top
            picture.Left = old_picture.Left
            old_picture.Delete()
            break
    else:
        # otherwise place the picture below the calling cell.
        top_left = sheet.Cells(caller.rect.last_row+2, caller.rect.last_col+1)
        picture.Top = top_left.Top
        picture.Left = top_left.Left

    # set the name of the new picture so we can find it next time
    picture.Name = figname

    # Delete the temporary file after the function has returned and
    # Excel has finished processing the new Picture object.
    async_call(os.unlink, filename)

    return "[Plotted '%s']" % figname


def xl_app():
    xl_window = get_active_object()
    xl_app = win32com.client.Dispatch(xl_window).Application
    return xl_app
