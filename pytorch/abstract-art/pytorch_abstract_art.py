from pyxll import xl_func, xl_app, async_call
import matplotlib.pyplot as plt
import win32gui, win32ui, win32con
import torch.nn as nn
import numpy as np
import tempfile
import torch
import os


class NN(nn.Module):
    """Neural network with Sequential layers"""

    def __init__(self, layers):
        super(NN, self).__init__()
        print(layers)
        self.layers = nn.Sequential(*layers)

    def forward(self, x):
        return self.layers(x)


@xl_func
def nn_Linear(in_features: int, out_features: int, bias: bool=True):
    """Create a linear transformation layer."""
    return nn.Linear(in_features, out_features, bias)


@xl_func
def nn_Tanh():
    """Create a Tanh transformation layer."""
    return nn.Tanh()


@xl_func
def nn_Sigmoid():
    """Create a Sigmoid transformation layer."""
    return nn.Sigmoid()


@xl_func("object[]: object")
def nn_Sequential(layers):
    """Create a neural network from a list of layers."""
    # remove any empty cells passed in
    layers = [layer for layer in layers if layer is not None]
    return NN(layers)


@xl_func
def nn_Run(net, image_name, scale=1, offset=-0.5, seed=None):
    """Run the neural network with random weights"""

    # See the seed for the RNG
    seed = int(seed) if seed is not None else torch.random.initial_seed()
    torch.manual_seed(seed)

    # Initialize the weights
    def init_weights(m):
        if isinstance(m, nn.Linear):
            nn.init.normal_(m.weight)

    net.apply(init_weights)

    # Find the Excel image
    xl = xl_app()
    sheet = xl.Caller.Worksheet
    image = sheet.Pictures(image_name)

    # Get the image size in pixels
    size_x, size_y = get_image_size(image)

    # Create the inputs
    inputs = np.zeros((size_y, size_x, 2))
    for x in np.arange(0, size_x, 1):
        for y in np.arange(0, size_y, 1):
            scaled_x = scale * ((float(x) / size_x) + offset)
            scaled_y = scale * ((float(y) / size_y) + offset)
            inputs[y][x] = np.array([scaled_x, scaled_y])

    inputs = inputs.reshape(size_x * size_y, 2)

    # Compute the results
    result = net(torch.tensor(inputs).type(torch.FloatTensor)).detach().numpy()
    result = result.reshape((size_y, size_x, 3))

    # Create a temporary file to write the result to
    file = create_temporary_file(suffix=".png")

    # Write the image to the file
    plt.imsave(file, result)
    file.flush()

    # Replace the old image with the new one
    new_image = sheet.Shapes.AddPicture(Filename=file.name,
                                        LinkToFile=0,  # msoFalse
                                        SaveWithDocument=-1,  # msoTrue
                                        Left=image.Left,
                                        Top=image.Top,
                                        Width=image.Width,
                                        Height=image.Height)

    image_name = image.Name
    image.Delete()
    new_image.Name = image_name

    return f"[{new_image.Name}]"


def get_image_size(image):
    """Return the size of an image in pixels as (width, height)."""
    # Get the size of the input image in pixels (Excel sizes are in points,
    # or 1/72th of an inch.
    xl = xl_app()
    dc = win32gui.GetDC(xl.Hwnd)
    pixels_per_inch_x = win32ui.GetDeviceCaps(dc, win32con.LOGPIXELSX)
    pixels_per_inch_y = win32ui.GetDeviceCaps(dc, win32con.LOGPIXELSY)
    size_x = int(image.Width * pixels_per_inch_x / 72)
    size_y = int(image.Height * pixels_per_inch_y / 72)

    return size_x, size_y


def create_temporary_file(suffix=None):
    """Create a named temporary file that is deleted automatically."""
    # Rather than delete the file when it's closed we delete it when the
    # windows message loop next runs. This gives Excel enough time to
    # load the image before the file disappears.
    file = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)

    def try_delete(filename):
        try:
            os.unlink(filename)
        except PermissionError:
            # retry if Excel is still accessing it
            async_call(os.unlink, filename)

    # Make sure the file gets deleted after the function is complete
    async_call(try_delete, file.name)

    return file
