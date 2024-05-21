"""
Example code showing how to draw a matplotlib figure embedded
in an Excel worksheet.

Matplotlib is used to plot a chart to an image, which is then
displayed as a Picture object in Excel.
"""
from pyxll import xl_func, plot
import pandas as pd
from matplotlib import pyplot as plt


@xl_func("numpy_column<float> xs, numpy_column<float> ys, int span: object")
def mpl_plot_ewma(xs, ys, span):
    # create the figure and axes for the plot
    fig, ax = plt.subplots()

    # calculate the moving average
    moving_average = pd.Series(ys, index=xs).ewm(span=span).mean()

    # plot the data
    ax.plot(xs, ys, alpha=0.4, label="Raw")
    ax.plot(xs, moving_average.values, label="EWMA")
    ax.legend()

    # Show the figure in Excel
    plot(fig)

    # Return the figure as an object
    return fig
