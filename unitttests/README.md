# Unit Testing PyXLL Functions

PyXLL provides an implementation of the `pyxll` module that can
be imported *outside* of Excel. This is so that you can import
your PyXLL code without it running in Excel, for unit-testing
as one example.

It is supplied as a wheel file in the downloaded PyXLL zip file
and can be installed using *pip*.

You will need to substitute the folder and wheel filename to where
you have unzipped PyXLL in the below example.

```bash
cd c:\Users\jondoe\pyxll
pip install  pyxll-4.4.2-cp37-none-win32.whl
```

The files in this folder demonstrate how the Python unittest
package can be used to test PyXLL code.
