"""
Custom excel types for pandas objects (eg dataframes).

For information about custom types in PyXLL see:
https://www.pyxll.com/docs/udfs.html#custom-types

For information about pandas see:
http://pandas.pydata.org/

Example code using 'pandastypes.py'.
"""

from pyxll import xl_func
import pandas as pa
import numpy as np


@xl_func("int rows, int cols: dataframe")
def make_random_dataframe(rows, cols):
    # create a random dataframe
    df_dict = {}
    for col_index in range(cols):
        col_name = chr(col_index + ord('A'))
        col_values = [np.random.random() for r in range(rows)]
        df_dict[col_name] = col_values

    df = pa.DataFrame(df_dict)

    # return it. The custom type will convert this to a 2d array that
    # excel will understand when this function is called as an array
    # function.
    return df


@xl_func("dataframe df, string col: float")
def sum_column(df, col):
    """take a dataframe and return the sum of a single column"""
    return df[col].sum()


@xl_func("dataframe df, string group_by_col, string value_col: series")
def group_by(df, group_by_col, value_col):
    """take a dataframe and group it by a single column and return the sum of another"""
    return df.groupby([group_by_col])[value_col].sum()

