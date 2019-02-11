"""
Highstock Demos
Two panes, candlestick and volume: http://www.highcharts.com/stock/demo/candlestick-and-volume

This code accompanies the blog post
https://www.pyxll.com/blog/interactive-charts-in-excel-with-highcharts
"""
from pyxll import xl_func
from highcharts import Highstock
from highcharts_xl import hc_plot


@xl_func
def hc_candlestick_and_volume(data, title, theme=None):
    H = Highstock()

    ohlc = []
    volume = []
    groupingUnits = [
        ['week', [1]],
        ['month', [1, 2, 3, 4, 6]]
    ]

    for i in range(len(data)):
        ohlc.append(
            [
            data[i][0], # the date
            data[i][1], # open
            data[i][2], # high
            data[i][3], # low
            data[i][4]  # close
            ]
            )
        volume.append(
            [
            data[i][0], # the date
            data[i][5]  # the volume
            ]
        )

    options = {
        'rangeSelector': {
            'selected': 1
        },

        'title': {
            'text': title
        },

        'yAxis': [{
            'labels': {
                'align': 'right',
                'x': -3
            },
            'title': {
                'text': 'OHLC'
            },
            'height': '60%',
            'lineWidth': 2
        }, {
            'labels': {
                'align': 'right',
                'x': -3
            },
            'title': {
                'text': 'Volume'
            },
            'top': '65%',
            'height': '35%',
            'offset': 0,
            'lineWidth': 2
        }],
    }

    H.add_data_set(ohlc, 'candlestick', 'OHLC', dataGrouping = {
        'units': groupingUnits
    })

    H.add_data_set(volume, 'column', 'Volume', yAxis = 1, dataGrouping = {
        'units': groupingUnits
    })

    H.set_dict_options(options)

    return hc_plot(H, title, theme)
