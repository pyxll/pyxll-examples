"""
Custom cell formatting example.

This example shows how to write a custom formatter that copies the
cell formatting for an array function when the result of the 
array function expands.

This can be useful where the Excel user applies their own formatting
to a range, but that range can change shape as the inputs to the
function change.

As array data often contains cell headers, the first n rows are
treated separately from the rest of the range to allow column
headers to be formatted differently from the rest of the data.

This example requires the 3rd party package pywin32 to be installed.

This code accompanies the blog post
https://www.pyxll.com/blog/custom-fill-formatter
"""
from pyxll import Formatter, xl_func
from win32com.client import constants


class FillFormatter(Formatter):
    """FillFormatter is a custom cell formatter used to apply
    formatting to data returned from an Excel array function.

    To use a formatter, pass it to the @xl_func decorator using
    the 'formatter' kwarg, for example::

        fill_formatter = FillFormatter()

        @xl_func(formatter=fill_formatter)
        def your_function(...):
            return an_array
    """
    
    def __init__(self, num_headers=1):
        """FillFormatter constructor.

        :param num_headers: Number of column header rows to treat
                            separately from the rest of the data
                            when formatting.
        """
        super().__init__()
        self.__num_headers = num_headers
        self.__prev_cell = None
    
    def clear(self, cell):
        """The clear method is called by PyXLL as part of the formatting.
        It is always immediately followed by a call to the 'apply' method.

        Here, rather than clear any formatting we store the cell that should
        be cleared to use in the apply method.
        """
        # Don't clear the previous cell(s) yet
        self.__prev_cell = cell

    
    def apply(self, cell, *args, **kwargs):
        """The apply method is called immediately after the clear method
        has been called.

        It's here were we can apply any custom formatting to the cell(s) that
        the array function has been returned to.

        We use the previously stored 'prev_cell' to copy the formatting from
        the previous range to the new range.
        """
        if self.__prev_cell is None:
            return
        
        # Reset prev_cell ahead of any future calls.
        prev_cell = self.__prev_cell
        self.__prev_cell = None

        # Get the previous and new Range objects from the cells as win32com objects.
        prev_range = prev_cell.to_range(com_package="win32com")
        new_range = cell.to_range(com_package="win32com")
        
        # Check if the range is shrinking and clear the formatting if it is.
        if prev_range.Rows.Count > new_range.Rows.Count:
            num_rows = prev_range.Rows.Count - new_range.Rows.Count
            rows = prev_range.GetOffset(RowOffset=new_range.Rows.Count)
            rows = rows.GetResize(RowSize=num_rows)
            rows.ClearFormats()
        
        if prev_range.Columns.Count > new_range.Columns.Count:
            num_cols = prev_range.Columns.Count - new_range.Columns.Count
            cols = prev_range.GetOffset(ColumnOffset=new_range.Columns.Count)
            cols = cols.GetResize(ColumnSize=num_cols)
            cols.ClearFormats()
        
        # Copy the formatting if the range has expanded.
        if new_range.Columns.Count > prev_range.Columns.Count \
        or new_range.Rows.Count > prev_range.Rows.Count:
            prev_rows = prev_range
            new_rows = new_range

            # If we have header rows then copy those separately from the rest.
            if self.__num_headers > 0:
                prev_header = prev_range.GetResize(RowSize=self.__num_headers)
                new_header = new_range.GetResize(RowSize=self.__num_headers)
                
                prev_header.Copy()
                new_header.PasteSpecial(Paste=constants.xlPasteFormats,
                                        Operation=constants.xlNone)
        
                if prev_range.Rows.Count > self.__num_headers:
                    prev_rows = prev_rows.GetOffset(RowOffset=self.__num_headers)
                    prev_rows = prev_rows.GetResize(RowSize=prev_range.Rows.Count - self.__num_headers)
                    
                    new_rows = new_rows.GetOffset(RowOffset=self.__num_headers)
                    new_rows = new_rows.GetResize(RowSize=new_range.Rows.Count - self.__num_headers)
        
            # Copy for the formatting from the previous range to the new one.
            prev_rows.Copy()
            new_rows.PasteSpecial(Paste=constants.xlPasteFormats,
                                  Operation=constants.xlNone)
        
            # End copy/paste mode.
            new_rows.Application.CutCopyMode = False


@xl_func("int rows, int columns: var[][]", auto_resize=True, formatter=FillFormatter())
def fill_formatter_example(rows, columns):
    """Sample function that simply returns a table of data for
    the purpose of demonstrating the FillFormatter class.
    """
    header = [f"COL{i+1}" for i in range(columns)]
    rows = [[c + r * columns for c in range(columns)] for r in range(rows)]
    return [header] + rows
