"""
Example code showing how to draw a matplotlib figure embedded
in an Excel worksheet.

Matplotlib is used to plot a chart to an image, which is then
displayed as a Picture object in Excel.
"""
from pyxll import xl_func, xlfCaller
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
    fig = Figure(figsize=(8, 6), dpi=75, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
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

    # if a picture with the same figname already exists then get the position
    # and size from the old picture and delete it.
    for old_picture in sheet.Pictures():
        if old_picture.Name == figname:
            height = old_picture.Height
            width = old_picture.Width
            top = old_picture.Top
            left = old_picture.Left
            old_picture.Delete()
            break
    else:
        # otherwise place the picture below the calling cell.
        top_left = sheet.Cells(caller.rect.last_row+2, caller.rect.last_col+1)
        top = top_left.Top
        left = top_left.Left
        width, height = fig.bbox.bounds[2:]

    # insert the picture
    # Ref: http://msdn.microsoft.com/en-us/library/office/ff198302%28v=office.15%29.aspx
    picture = sheet.Shapes.AddPicture(Filename=filename,
                                      LinkToFile=0,  # msoFalse
                                      SaveWithDocument=-1,  # msoTrue
                                      Left=left,
                                      Top=top,
                                      Width=width,
                                      Height=height)

    # set the name of the new picture so we can find it next time
    picture.Name = figname

    # delete the temporary file
    os.unlink(filename)

    return "[Plotted '%s']" % figname


def xl_app():
    xl_window = get_active_object()
    xl_app = win32com.client.Dispatch(xl_window).Application
    return xl_app
