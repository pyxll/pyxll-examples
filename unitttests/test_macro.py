"""This example shows using the unittest.mock package
to mock the results of xl_app.Range.

Note: unitest.mock was added in Python 3.3. If you are
using an earlier version then you will need to pip
install 'mock' and change the import to import from
'mock' instead of 'unittest.mock'.

If you cannot import pyxll then see the notes in the README.md
file alongside this file.
"""
from pyxll import xl_macro, xl_app
from unittest.mock import patch, MagicMock
import unittest


#
# This is an example macro that we might want to test.
#
# It is just a simple example that uses PyXLL's
# xl_app function so we can demonstrate how to mock
# the Excel.Application object returned by xl_app.
#
@xl_macro("Example Macro")
def example_macro():
    # Get the Excel.Application object
    xl = xl_app()

    # Add A1 to A2 and store the result in A3
    a = xl.Range("A1")
    b = xl.Range("A2")
    c = xl.Range("A3")
    c.Value = a.Value + b.Value


class MacroTests(unittest.TestCase):

    # IMPORTANT
    #
    # When mocking we must patch *where an object is looked up*.
    # This is not necessarily the same place as where it is defined.
    # Here our code uses the local module variable 'xl_app' and
    # *not' 'pyxll.xl_app' and so we have to patch the local 'xl_app'.
    #
    # You will want to replace this with @patch("your-module.xl_app")
    # if you are using "import xl_app from pyxll", or
    # @patch("pyxll.xl_app") if you are using "import pyxll".
    #
    @patch(f"{__name__}.xl_app")
    def test_macro(self, mock_xl_app):
        # Mock the "Range" property with a side effect to return a different
        # mocked 'Range' object for different addresses, with "Value" set
        # for address A1 and A2.
        mocked_range_values = {
            "A1": 1.0,
            "A2": 2.0
        }

        mocked_ranges = {}

        # This function is the side_effect of xl_app().Range and
        # returns a different mock object for each address.
        def create_mock_range(address):
            if address in mocked_ranges:
                return mocked_ranges[address]

            mocked_range = MagicMock()
            if address in mocked_range_values:
                mocked_range.Value = mocked_range_values[address]

            mocked_ranges[address] = mocked_range
            return mocked_range

        mock_xl = mock_xl_app()
        mock_xl.Range.side_effect = create_mock_range

        # Call the example macro
        example_macro()

        # Check that A3 has been set to 3.0 (A3 = A1 + A2; A1 = 1.0; A2 = 2.0)
        self.assertEqual(3.0, mock_xl_app().Range("A3").Value)
