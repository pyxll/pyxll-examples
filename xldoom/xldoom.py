"""
DOOM in Excel

Runs Doom and displays frames in Excel using an RTD (Real Time Data)
function and a custom cell formatter.

See https://youtu.be/m3K7wUTX-SY
"""
from pyxll import xl_func, Formatter
from pathlib import Path
import cydoomgeneric as cdg
import scipy.ndimage
import numpy as np
import threading
import itertools
import time


_frame = None
_frame_event = threading.Event()
_thread = None
_thread_lock = threading.Lock()


def _draw_frame(pixels):
    """Called by the Doom main loop whenever a frame is drawn."""
    # Update the global frame and let any other threads know the data is ready
    global _frame
    _frame = pixels
    _frame_event.set()
    time.sleep(0.01)


def _get_key():
    """Called by the Doom main loop to get inputs."""
    # We don't handle any input
    pass


def _start_doom():
    """Starts Doom in a background thread, if not already started.
    """
    global _thread

    # Acquire the lock before updating global variable _thread
    with _thread_lock:
        # If the doom thread isn't alrady running then start Doom in its own background thread
        if _thread is None:
            # Function to initialize and run Doom in our background thread
            def doom_thread_func(wadfile):
                cdg.init(640, 400, _draw_frame, _get_key)
                cdg.main(["", "-iwad", str(wadfile), "-playdemo"])

            # Check the wad file exists
            wadfile = Path(__file__).parent / "Doom1.WAD"
            if not wadfile.exists():
                raise RuntimeError(f"WAD file '{wadfile}' not found.")

            _thread = threading.Thread(target=doom_thread_func, args=(wadfile,))
            _thread.daemon = True
            _thread.start()

    return True


class DoomFormatter(Formatter):
    """Formatter class that set the cell interior color to the 
    cell value (colour value as uint32 BGR) and sets the number
    format to blank.
    """

    def apply(self, cell, value=None, **kwargs):
        if not isinstance(value, np.ndarray):
            return

        # This style displays an empty cell instead of the cell value
        self.apply_style(cell, {
            "number_format": ";;;"
        })

        # Iterate over each cell in the entire range, setting the
        # interior color to be the value from the numpy array.
        cell = cell.resize(1, 1)
        for y in range(0, value.shape[0]):
            for x in range(0, value.shape[1]):
                self.apply_style(cell.offset(y, x), {
                    "interior_color": value[y, x]
                })


@xl_func("float scale: rtd<union<str, numpy_array>>", formatter=DoomFormatter())
def doom(scale=0.15):
    # Ensure the Doom background thread is running
    _start_doom()

    yield "Please wait..."

    # Yield pixel data as it becomes available
    while True:
        _frame_event.wait()

        # Scale the frame buffer and yield
        pixels = scipy.ndimage.zoom(_frame, [scale, scale, 1])

        # pixels is a 2d array of [b, g, r, a] values.
        # Select the (b, g, r) part and convert to uint32 #bgr values
        pixels = pixels.astype(np.uint32)

        # Get the blue, green and red channels
        blue = pixels[:,:,0]
        green = pixels[:,:,1]
        red = pixels[:,:,2]

        # Quantize to limit the number of colors
        blue = (blue >> 3) << 3
        green = (green >> 3) << 3
        red = (red >> 3) << 3

        # Combine into #bgr values
        pixels = (blue << 16) | (green << 8) | red
        yield pixels


if __name__ == "__main__":
    # Print a few frames to show it working
    for pixels in itertools.islice(doom(), 25):
        print(pixels)
