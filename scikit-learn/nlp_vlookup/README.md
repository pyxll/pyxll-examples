# NLP.VLOOKUP

A more flexible VLOOKUP using Natural Language Processing
from scikit-learn.

To use in Excel add the folder this file is in to your pythonpath
in the pyxll.cfg file and add 'nlp_vlookup' to the modules list
in the same file.

Once installed you may use the "NLP.VLOOKUP" function in Excel

```
=NLP.VLOOKUP(
    lookup_value,
    table_array,
    col_index_num, (optional)
    include_score, (default is False)
    all_matches, (default is False)
    threshold (default is 0.5)
)
```

The NLP.VLOOKUP function works in a similar way to VLOOKUP but can match
similar words not just identical words.

There is an example workbook in this folder that demonstrates how to use
the NLP.VLOOKUP function.
