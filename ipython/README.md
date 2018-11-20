# IPython Excel Integration

These samples run an IPython kernel in-process in Excel and launch an IPython Qt terminal connected to that kernel.

Some of the examples also include a ribbon.xml file which can either be added to your pyxll.cfg file or combined with any existing
ribbon xml file you might have.

Add the one corresponding to your version of IPython to your pyxll.cfg file, e.g.

```ini
[PYTHON]
pythonpath =
  {path to pyxll-examples}/ipython/7.x

[PYXLL]
modules =
  ipython
ribbon = {path to pyxll-examples}/ipython/7.x/ribbon.xml
```

See requirements.txt in the version sub-folders for dependencies.
