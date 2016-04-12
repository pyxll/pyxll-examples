"""
Start an IPython Qt console connected to the python session running in Excel.

This doesn't work with an IPython notebook as it's not possible to connect
a notebook to an existing kernel, the notebook app always creates its own.

This version is intended to work with IPython versions 1.x only.
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
    global _ipython_app, _kernel_running

    if _ipython_app and _kernel_running:
        return _ipython_app

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
            global _kernel_running

            # if the kernel has been closed then run the event loop until it gets to the
            # stop event added by IPKernelApp.shutdown_request
            if self.kernel.shell.exit_now:
                _log.debug("IPython kernel stopping (%s)" % self.connection_file)
                timer.kill_timer(timer_id)
                loop.start()
                _kernel_running = False
                return

            # otherwise call the event loop but stop immediately if there are no pending events
            loop.add_timeout(0, lambda: loop.add_callback(loop.stop))
            loop.start()

        global _kernel_running
        _kernel_running = True
        timer.set_timer(100, poll_ioloop)

    IPKernelApp.start = _IPKernelApp_start

    # IPython expects sys.__stdout__ to be set
    sys.__stdout__ = sys.stdout
    sys.__stderr__ = sys.stderr

    # call the API embed function, which will use the monkey-patched method above
    IPython.embed_kernel()

    _ipython_app = IPKernelApp.instance()

    # patch ipapp so anything else trying to get a terminal app (e.g. ipdb)
    # gets our IPKernalApp.
    from IPython.terminal.ipapp import TerminalIPythonApp
    TerminalIPythonApp.instance = lambda: _ipython_app
    __builtins__["get_ipython"] = lambda: _ipython_app.shell

    return _ipython_app


def _launch_qt_console(ppid, connection_file):
    """called as a new process"""
    from IPython.terminal.ipapp import TerminalIPythonApp
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

