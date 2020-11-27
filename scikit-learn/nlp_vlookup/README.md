# NLP.VLOOKUP

A more flexible VLOOKUP using Natural Language Processing
from scikit-learn.

The following video accompanies this code:

[![Improving VLOOKUP with NLP Video](http://img.youtube.com/vi/bqMH62QniC8/0.jpg)](https://www.youtube.com/watch?v=bqMH62QniC8 "Improving VLOOKUP with NLP Video")

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
