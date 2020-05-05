"""This example shows a unit test of a simple Excel
worksheet function.

If you cannot import pyxll then see the notes in the README.md
file alongside this file.
"""
from pyxll import xl_func
import unittest


#
# This is an example UDF that we might want to test.
# For a more complex example showing mocking, see the
# test_macro.py test in the same folder as this file.
#
@xl_func
def example_udf(a, b, c):
    return (a + b) * c


class UdfTests(unittest.TestCase):

    def test_udf(self):
        # When called outside of Excel the @xl_func decorator doesn't
        # do anything so we can simply call the function in the same
        # way as any other Python function.
        actual = example_udf(1, 2, 3)
        expected = (1 + 2) * 3

        self.assertEqual(expected, actual)
