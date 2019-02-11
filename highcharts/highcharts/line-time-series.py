"""
Highcharts Demos
Time series, zoomable: http://www.highcharts.com/demo/line-time-series

This code accompanies the blog post
https://www.pyxll.com/blog/interactive-charts-in-excel-with-highcharts
"""
from pyxll import xl_func, get_type_converter
from highcharts_xl import hc_plot
from highcharts import Highchart


@xl_func
def hc_line_time_series(data, title, y_title=None, theme=None, xl_dates=False):
    """Plots a 2d array of times and data points."""
    # Convert Excel dates to datetimes if necessary
    if xl_dates:
        to_date = get_type_converter("var", "datetime")
        data = [(to_date(d), v) for d, v in data]

    H = Highchart()

    H.set_options('chart', {
        'zoomType': 'x'
    })

    H.set_options('xAxis', {
        'type': 'datetime'
    })

    H.set_options('yAxis', {
        'title': {
            'text': y_title or "Values"
        }
    })

    H.set_options('title', {
        'text': title or "Time Series"
    })

    H.set_options('legend', {
        'enabled': False
    })

    H.add_data_set(data, 'area')

    H.set_options('plotOptions', {
        'area': {
            'fillColor': {
                'linearGradient': { 'x1': 0, 'y1': 0, 'x2': 0, 'y2': 1},
                'stops': [
                    [0, "Highcharts.getOptions().colors[0]"],
                    [1, "Highcharts.Color(Highcharts.getOptions().colors[0]).setOpacity(0).get('rgba')"]
                ]},
            'marker': {
                'radius': 2
            },
            'lineWidth': 1,
            'states': {
                'hover': {
                    'lineWidth': 1
                }
            },
            'threshold': None
        }
    })

    return hc_plot(H, title, theme)
\