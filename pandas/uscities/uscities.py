"""
This module contains the Excel functions used in the uscities.xlsx workbook.

See README.md for more information about this example.
"""
from pyxll import xl_func, xl_arg, xl_return
import pandas as pd
import os

# The signature specifies that 'path' is a string, and the result will be
# returned to Excel as an object handle instead of expanding the DataFrame
# to a range in Excel.
@xl_func("string path: object")
def load_csv_data(path):
    """Loads a CSV file as a pandas DataFrame."""

    # If the path is a relative path make it absolute from this folder.
    if path.startswith("."):
        path = os.path.join(os.path.dirname(__file__), path)

    # Load the CSV file and return it as a pandas DataFrame
    df = pd.read_csv(path)
    return df


# Here the signature specifies that the returned DataFrame should be be
# returned to Excel as a DataFrame (rather than an object as above) and
# it should include the index.
@xl_func("dataframe: dataframe<index=True>")
def df_describe(df):
    """Get some stats about the dataframe"""
    return df.describe()


@xl_func("dataframe df, var[][] filter_values: object")
def df_filter(df, filter_values):
    """
    Filter a dataframe and return the filtered result.
    Takes a dataframe and a list of (boolean operator, column, operation, value)
    rows, eg
        [ (None,  "A", "==", "X"),
          ("AND", "B", "==", "Y"),
          ("OR",  "C", "!=", "Z") ]
    """
    mask = None
    mask_op = "AND"

    for bool_op, col, row_op, value in filter_values:
        # The 'bool_op' variable can be None, in which case the previous value is used.
        if bool_op is not None:
            mask_op = bool_op.upper()

        # Skip blank rows
        if not col:
            continue

        # Get the values that match
        if row_op == "==":
            row_mask = df[col] == value
        elif row_op == "!=":
            row_mask = df[col] != value
        elif row_op == ">=":
            row_mask = df[col] >= value
        elif row_op == ">":
            row_mask = df[col] > value
        elif row_op == "<=":
            row_mask = df[col] <= value
        elif row_op == "<":
            row_mask = df[col] < value
        else:
            raise ValueError(f"Unexpected operator '{op}'")

        # If it's the first row then there's nothing to combine it with
        if mask is None:
            mask = row_mask
            continue

        # Combine the mask from this row with the main mask
        if mask_op == "AND":
            mask = mask & row_mask
        elif mask_op == "OR":
            mask = mask | row_mask
        else:
            raise ValueError(f"Unexpected mask operator '{mask_op}'")

    # Filter the DataFrame using the mask
    if mask is not None:
        df = df[mask]

    return df


@xl_func("dataframe df, int n: dataframe<index=True>")
def df_head(df, n):
    """Get the first n rows of a dataframe"""
    return df.head(n)


# I've used xl_arg and xl_return instead of a signature here. The two methods
# are equivalent, but using xl_arg and xl_return can be easier when dealing
# with a number of complex arguments.
@xl_func
@xl_arg("df", "dataframe")
@xl_arg("columns", "string[]")
@xl_arg("agg_funcs", "dict<string, string>")
@xl_arg("transpose", "bool")
@xl_return("dataframe", index=True, multi_sparse=False)
def df_pivot_table(df, columns, agg_funcs, transpose=False):
    # remove empty columns
    columns = filter(None, columns)

    # pivot the data
    df = pd.pivot_table(df, columns=columns, aggfunc=agg_funcs)

    # transpose if needed
    if transpose:
        df = df.transpose()

    return df
