"""
Highcharts Demos
Basic line: https://www.highcharts.com/demo/line-basic

This code accompanies the blog post
https://www.pyxll.com/blog/interactive-charts-in-excel-with-highcharts
"""
from pyxll import xl_func
from highcharts_xl import hc_plot
from highcharts import Highchart


@xl_func("var[][] data, str title, str subtitle, str y_axis, str[] labels, str theme")
def hc_basic_line(data, title, subtitle=None, y_axis=None, labels=None, theme=None):
    H = Highchart()

    H.set_options("title", {
        "text": title
    })

    if subtitle:
        H.set_options("subtitle", {
            "text": subtitle
        })

    if y_axis:
        H.set_options("yAxis", {
            "title": {
                "text": y_axis
            }
        })

    H.set_options("legend", {
        "layout": "vertical",
        "align": "right",
        "verticalAlign": "middle"
    })

    # transform the data from a list of rows to a list of columns
    data_t = list(zip(*data))

    # Use the first column as the X axis
    x_axis = data_t[0]

    H.set_options("xAxis", {
        "categories": x_axis,
        "tickmarkPlacement": "on"
    })

    # And the remaining columns as the graph data
    if not labels:
        labels = ["series_%d" % i for i in range(len(data_t)-1)]

    for label, series in zip(labels, data_t[1:]):
        H.add_data_set(series, series_type='line', name=label)

    return hc_plot(H, title, theme)
