"""
Start an IPython Qt console connected to the python session running in Excel.

This doesn't work with an IPython notebook as it's not possible to connect
a notebook to an existing kernel, the notebook app always creates its own.

This version is intended to work with IPython 0.13 only.
"""
from pyxll import xl_menu
import logging
import timer
import sys
import os

_log = logging.getLogger(__name__)

_kernel_running = False
_ipython_app = None    


@xl_menu("IPython")
def ipython_qtconsole():
    """
    Launches an IPython Qt console
    """
    # try to set sys.executable if it's not already set
    _fixup_sys_executable()

    # start the IPython kernel
    app = _start_kernel()

    # start a subprocess to run the Qt console
    from multiprocessing import Process    
    proc = Process(target=_launch_qt_console, args=[os.getpid(), app.connection_file])
    proc.daemon = True
    proc.start()


def _fixup_sys_executable():
    """
    Set sys.executable to the default python executable, if it's not already set.
    This expects that python will be installed as the default python and pythonw.exe
    exists on the PATH.

    If you get errors when trying to launch the Qt IPython prompt with multiprocessing
    check this, and set sys.executable to the absolute location of your installed python.
    """
    # don't do anything if it's already set
    if sys.executable and os.path.basename(sys.executable) in ("python.exe", "pythonw.exe"):
        return

    executable = _which("pythonw.exe")
    if not executable:
        _log.error("Couldn't find pythonw.exe on the PATH. Starting the subprocess will fail.")
        return

    _log.info("Setting sys.executable to '%s'" % executable)
    sys.executable = executable


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
    from IPython.zmq.ipkernel import IPKernelApp
    from zmq.eventloop import ioloop

    global _kernel_running, _ipython_app
    if _kernel_running:
        return _ipython_app

    # get the app if it exists, or set it up if it doesn't
    if IPKernelApp.initialized():
        app = IPKernelApp.instance()
    else:
        app = IPKernelApp.instance()
        app.initialize()

        # Undo unnecessary sys module mangling from init_sys_modules.
        # This would not be necessary if we could prevent it
        # in the first place by using a different InteractiveShell
        # subclass, as in the regular embed case.
        main = app.kernel.shell._orig_sys_modules_main_mod
        if main is not None:
            sys.modules[app.kernel.shell._orig_sys_modules_main_name] = main

    app.kernel.user_module = sys.modules[__name__]
    app.kernel.user_ns = {}

    # patch in auto-completion support
    # added in https://github.com/ipython/ipython/commit/f4be28f06c2b23cd8e4a3653b9e84bde593e4c86
    # we effectively make the same patches via monkeypatching
    from IPython.core.interactiveshell import InteractiveShell
    old_set_completer_frame = InteractiveShell.set_completer_frame

    # restore old set_completer_frame that gets no-op'd out in ZmqInteractiveShell.__init__
    bound_scf = old_set_completer_frame.__get__(app.shell, InteractiveShell)
    app.shell.set_completer_frame = bound_scf
    app.shell.set_completer_frame()

    # start the kernel
    app.kernel.start()

    # set up a timer to periodically poll the zmq ioloop
    loop = ioloop.IOLoop.instance()

    def poll_ioloop(timer_id, time):
        global _kernel_running

        # if the kernel has been closed then run the event loop until it gets to the
        # stop event added by IPKernelApp.shutdown_request
        if app.kernel.shell.exit_now:
            _log.debug("IPython kernel stopping (%s)" % app.connection_file)
            timer.kill_timer(timer_id)
            ioloop.IOLoop.instance().start()
            _kernel_running = False
            return

        # otherwise call the event loop but stop immediately if there are no pending events
        loop.add_timeout(0, lambda: loop.add_callback(loop.stop))
        ioloop.IOLoop.instance().start()

    _log.debug("IPython kernel starting. Use '--existing %s' to connect." % app.connection_file)
    timer.set_timer(100, poll_ioloop)
    _kernel_running = True
    
    _ipython_app = app
    return _ipython_app


def _launch_qt_console(ppid, connection_file):
    """called as a new process"""
    from IPython.frontend.terminal.ipapp import TerminalIPythonApp
    import threading
    import psutil
    import time
    
    # start a thread to kill this process when the parent process exits
    def thread_func():
        while True:
            if not psutil.pid_exists(ppid):
                os._exit(1)
            time.sleep(5)
    thread = threading.Thread(target=thread_func)
    thread.daemon = True
    thread.start()
    
    # start the qtconsole app
    app = TerminalIPythonApp.instance()
    app.initialize(["qtconsole", "--existing", connection_file])
    app.start()
