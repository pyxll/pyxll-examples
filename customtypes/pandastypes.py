"""
Custom excel types for pandas objects (eg dataframes).

For information about custom types in PyXLL see:
https://www.pyxll.com/docs/udfs.html#custom-types

For information about pandas see:
http://pandas.pydata.org/

Including this module in your pyxll config adds the following custom types that can
be used as return types to your pyxll functions:

	- dataframe
	- series
	- series_t

Dataframes with multi-index indexes or columns will be returned with the columns and
index values in the resulting array. For normal indexes, the index will only be
returned as part of the resulting array if the index is named.

eg::

	from pyxll import xl_func
	import pandas as pa

	@xl_func("int rows, int cols, float value: dataframe")
	def make_empty_dataframe(rows, cols, value):
		# create an empty dataframe
		df = pa.DataFrame({chr(c + ord('A')) : value for c in range(cols)}, index=range(rows))
		
		# return it. The custom type will convert this to a 2d array that
		# excel will understand when this function is called as an array
		# function.
		return df

In excel (use Ctrl+Shift+Enter to enter an array formula)::

	=make_empty_dataframe(3, 3, 100)
	
	>>  A	B	C
	>> 100	100	100
	>> 100	100	100
	>> 100	100	100

"""
from pyxll import xl_return_type
import pandas as pa
import numpy as np

@xl_return_type("dataframe", "var")
def _dataframe_to_var(df):
    """return a list of lists that excel can understand"""
    if not isinstance(df, pa.DataFrame):
        return df
    df = df.applymap(lambda x: RuntimeError() if isinstance(x, float) and np.isnan(x) else x)
 
    index_header = [str(df.index.name)] if df.index.name is not None else []
    if isinstance(df.index, pa.MultiIndex):
        index_header = [str(x) or "" for x in df.index.names]

    if isinstance(df.columns, pa.MultiIndex):
        result = [([""] * len(index_header)) + list(z) for z in zip(*list(df.columns))]
        for header in result:
            for i in range(1, len(header) - 1):
                if header[-i] == header[-i-1]:
                    header[-i] = ""

        if index_header:
            column_names = [x or "" for x in df.columns.names]
            for i, col_name in enumerate(column_names):
                result[i][len(index_header)-1] = col_name
    
            if column_names[-1]:
                index_header[-1] += (" \ " if index_header[-1] else "") + str(column_names[-1])

            num_levels = len(df.columns.levels)
            result[num_levels-1][:len(index_header)] = index_header
    else:
        if index_header and df.columns.name:
            index_header[-1] += (" \ " if index_header[-1] else "") + str(df.columns.name)
        result = [index_header + list(df.columns)]    

    if isinstance(df.index, pa.MultiIndex):
        prev_ix = None
        for ix, row in df.iterrows():
            header = list(ix)
            if prev_ix:
                header = [x if x != px else "" for (x, px) in zip(ix, prev_ix)]
            result.append(header + list(row))
            prev_ix = ix

    elif index_header:
        for ix, row in df.iterrows():
            result.append([ix] + list(row))
    else:
        for ix, row in df.iterrows():
            result.append(list(row))

    return result


@xl_return_type("series", "var")
def _series_to_var(s):
    """return a list of lists that excel can understand"""
    if not isinstance(s, pa.Series):
        return s
    s = s.apply(lambda x: RuntimeError() if isinstance(x, float) and np.isnan(x) else x)
    return list(map(list, s.items()))


@xl_return_type("series_t", "var")
def _series_to_var_transform(s):
    """return a list of lists that excel can understand"""
    if not isinstance(s, pa.Series):
        return s
    s = s.apply(lambda x: RuntimeError() if isinstance(x, float) and np.isnan(x) else x)
    return list(zip(*s.items()))
