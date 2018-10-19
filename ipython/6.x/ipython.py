"""
Start an IPython Qt console connected to the python session running in Excel.

This doesn't work with an IPython notebook as it's not possible to connect
a notebook to an existing kernel, the notebook app always creates its own.

This version is intended to work with IPython versions 6.x only.

This example requires sys.executable to be set, and so it's recommended
that the following is added to the pyxll.cfg file:

[PYTHON]
executable = <path to your python installation>/pythonw.exe

"""
from pyxll import xl_menu, xl_app, get_type_converter
import pyxll
import logging
import timer
import sys
import os

_log = logging.getLogger(__name__)

try:
    import win32api
except ImportError:
    win32api = None

if getattr(sys, "_ipython_kernel_running", None) is None:
    sys._ipython_kernel_running = False

if getattr(sys, "_ipython_app", None) is None:
    sys._ipython_app = False


@xl_menu("Open QtConsole", menu="IPython")
def ipython_qtconsole(*args):
    """
    Launches an IPython Qt console
    """
    try:
        # start the IPython kernel
        app = _start_kernel()

        # start a subprocess to run the Qt console
        # run jupyter in it's own process
        _launch_qt_console(app.connection_file)
    except:
        if win32api:
            win32api.MessageBox(None, "Error starting IPython Qt console")
        _log.error("Error starting IPython Qt console", exc_info=True)


@xl_menu("Selection to IPython", menu="IPython")
def set_selection_in_ipython(*args):
    """
    Gets the value of the selected cell and copies it to
    the globals dict in the IPython kernel.
    """
    try:
        if not getattr(sys, "_ipython_app", None) or not sys._ipython_kernel_running:
            raise Exception("IPython kernel not running")

        xl = xl_app(com_package="win32com")
        selection = xl.Selection
        if not selection:
            raise Exception("Nothing selected")

        value = selection.Value

        # convert any cached objects (PyXLL >= 4 only)
        pyxll_version = int(pyxll.__version__.split(".")[0])
        if pyxll_version >= 4 and isinstance(value, str):
            try:
                to_object = get_type_converter("var", "object")
                value = to_object(value)
            except KeyError:
                pass

        # set the value in the shell's locals
        sys._ipython_app.shell.user_ns["_"] = value
        print("\n\n>>> Selected value set as _")
    except:
        if win32api:
            win32api.MessageBox(None, "Error setting selection in Excel")
        _log.error("Error setting selection in Excel", exc_info=True)


def _which(program):
    """find an exe's full path by looking at the PATH environment variable"""
    def is_exe(fpath):
        return os.path.isfile(fpath) and os.access(fpath, os.X_OK)

    fpath, fname = os.path.split(program)
    if fpath:
        if is_exe(program):
            return program
    else:
        for path in os.environ["PATH"].split(os.pathsep):
            path = path.strip('"')
            exe_file = os.path.join(path, program)
            if is_exe(exe_file):
                return exe_file

    return None


def _start_kernel():
    """starts the ipython kernel and returns the ipython app"""
    if sys._ipython_app and sys._ipython_kernel_running:
        return sys._ipython_app

    import IPython
    from ipykernel.kernelapp import IPKernelApp
    from zmq.eventloop import ioloop

    # patch IPKernelApp.start so that it doesn't block
    def _IPKernelApp_start(self):
        if self.poller is not None:
            self.poller.start()
        self.kernel.start()

        # set up a timer to periodically poll the zmq ioloop
        loop = ioloop.IOLoop.instance()

        def poll_ioloop(timer_id, time):
            # if the kernel has been closed then run the event loop until it gets to the
            # stop event added by IPKernelApp.shutdown_request
            if self.kernel.shell.exit_now:
                _log.debug("IPython kernel stopping (%s)" % self.connection_file)
                timer.kill_timer(timer_id)
                loop.start()
                sys._ipython_kernel_running = False
                return

            # otherwise call the event loop but stop immediately if there are no pending events
            loop.add_timeout(0, lambda: loop.add_callback(loop.stop))
            loop.start()

        sys._ipython_kernel_running = True
        timer.set_timer(100, poll_ioloop)

    IPKernelApp.start = _IPKernelApp_start

    # IPython expects sys.__stdout__ to be set
    sys.__stdout__ = sys.stdout
    sys.__stderr__ = sys.stderr

    # call the API embed function, which will use the monkey-patched method above
    IPython.embed_kernel()

    ipy = IPKernelApp.instance()

    # Keep a reference to the kernel even if this module is reloaded
    sys._ipython_app = ipy

    # patch user_global_ns so that it always references the user_ns dict
    setattr(ipy.shell.__class__, 'user_global_ns', property(lambda self: self.user_ns))

    # patch ipapp so anything else trying to get a terminal app (e.g. ipdb) gets our IPKernalApp.
    from IPython.terminal.ipapp import TerminalIPythonApp
    TerminalIPythonApp.instance = lambda: ipy
    __builtins__["get_ipython"] = lambda: ipy.shell.__class__

    # Use the inline matplotlib backend
    mpl = ipy.shell.find_magic("matplotlib")
    if mpl:
        mpl("inline")

    return ipy


def _launch_qt_console(connection_file):
    """Starts the jupyter console"""
    from subprocess import Popen

    # Find juypter-qtconsole.exe in the Scripts path local to python.exe
    exe = None
    if sys.executable and os.path.basename(sys.executable) in ("python.exe", "pythonw.exe"):
        path = os.path.join(os.path.dirname(sys.executable), "Scripts")
        exe = os.path.join(path, "jupyter-qtconsole.exe")

    # If it wasn't found look for it on the system path
    if exe is None or not os.path.exists(exe):
        exe = _which("jupyter-qtconsole.exe")

    if exe is None or not os.path.exists(exe):
        raise Exception("jupyter-qtconsole.exe not found")

    # run jupyter in it's own process
    cmd = [exe, "--existing", connection_file]
    proc = Popen(cmd, shell=True)
    if proc.poll() is not None:
        raise Exception("Command '%s' failed to start" % " ".join(cmd))
