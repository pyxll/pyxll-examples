"""
Functions for display charts in Excel using Highcharts.

Please be aware that the Highcharts project itself, as well as Highmaps and Highstock,
are only free for non-commercial use under the Creative Commons Attribution-NonCommercial license.
Commercial use requires the purchase of a separate license. Pop over to Highcharts for more information.

This code accompanies the blog post
https://www.pyxll.com/blog/interactive-charts-in-excel-with-highcharts
"""
from pyxll import xl_func, xl_app, xlfCaller
from highcharts.highstock.highstock_helper import jsonp_loader
from bs4 import BeautifulSoup
import tempfile
import timer
import re
import os


def hc_plot(chart, control_name, theme=None):
    """
    This function is used by the other plotting functions to render the chart as html
    and display it in Excel.
    """
    # add the theme if there is one
    if theme:
        chart.add_JSsource(["https://code.highcharts.com/themes/%s.js" % theme])

    # get the calling sheet
    caller = xlfCaller()
    sheet_name = caller.sheet_name

    # split into workbook and sheet name
    match = re.match("^\[(.+?)\](.*)$", sheet_name.strip("'\""))
    if not match:
        raise Exception("Unexpected sheet name '%s'" % sheet_name)
    workbook, sheet = match.groups()

    # get the Worksheet object
    xl = xl_app()
    workbook = xl.Workbooks(workbook)
    sheet = workbook.Sheets(sheet)

    # find the existing webbrowser control, or create a new one
    try:
        control = sheet.OLEObjects(control_name[:31])
        browser = control.Object
    except:
        control = sheet.OLEObjects().Add(ClassType="Shell.Explorer.2",
                                         Left=147,
                                         Top=60.75,
                                         Width=400,
                                         Height=400)
        control.Name = control_name[:31]
        browser = control.Object

    # set the chart aspect ratio to match the browser
    if control.Width > control.Height:
        chart.set_options("chart", {
             "height": "%d%%" % (100. * control.Height / control.Width)
         })
    else:
        chart.set_options("chart", {
             "width": "%d%%" % (100. * control.Width / control.Height)
         })

    # get the html and add the 'X-UA-Compatible' meta-tag
    soup = BeautifulSoup(chart.htmlcontent)
    metatag = soup.new_tag("meta")
    metatag.attrs["http-equiv"] = "X-UA-Compatible"
    metatag.attrs['content'] = "IE=edge"
    soup.head.insert(0, metatag)

    # write out the html for the browser to render
    fh = tempfile.NamedTemporaryFile("wt", suffix=".html", delete=False)
    filename = fh.name

    # clean up the file after 10 seconds to give the browser time to load
    def on_timer(timer_id, time):
        timer.kill_timer(timer_id)
        os.unlink(filename)
    timer.set_timer(10000, on_timer)

    fh.write(soup.prettify())
    fh.close()

    # navigate to the temporary file
    browser.Navigate("file://%s" % filename)

    return "[%s]" % control_name


@xl_func("string: object")
def hc_load_sample_data(name):
    """Loads Highcharts sample data from https://www.highcharts.com/samples/data/jsonp.php."""
    url = "https://www.highcharts.com/samples/data/jsonp.php?filename=%s.json&callback=?" % name
    return jsonp_loader(url)


@xl_func("object: var", auto_resize=True, category="Highcharts")
def hc_explode(data):
    """Explode a data object into an array.
    Caution: This may result in returning a lot of data to Excel.
    """
    return data
