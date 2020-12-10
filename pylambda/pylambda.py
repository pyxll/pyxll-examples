"""
An Excel function for creating new named Excel functions from
Python expressions.

Excel functions are created by building a new Python function that
evaluates the expression, exposed to Excel using the xl_func
decorator.

The 'recalc_on_open' option is used so that the functions are
available as soon as the workbook containing them is opened with
the need to recalculate.
"""
from pyxll import xl_func


@xl_func("str name, str[] args, str expr, str return_type, str[] imports: str", recalc_on_open=True)
def pylambda(name, args, expr, return_type="var", imports=[]):
    """
    Create a new Excel function that evaluates a Python expression.

    :param name: Name of the new Excel function.
    :param args: List of arguments with optional types.
    :param expr: Python expression that will be the body of the function.
    :param return_type: Return type of the Excel function.
    :param imports: List of Python packages to import.
    """
    # Get the argument names and types
    arg_names = []
    arg_types = []
    for arg in args:
        arg_name, *arg_type = map(str.strip, arg.split(":", 1))
        arg_type = arg_type[0] if arg_type else "var"
        arg_names.append(arg_name)
        arg_types.append(arg_type)

    # Build a signature string using the argument types
    signature = ", ".join([f"{t} {n}" for n, t in zip(arg_names, arg_types)])
    signature += f": {return_type}"

    # Build the Python function as a string using the @xl_func decorator
    func = f"""
@xl_func("{signature}", name="{name}")
def func({', '.join(arg_names)}):
    return {expr}
"""

    # Create a namespace dictionary to exec the code in
    namespace = {
        "xl_func": xl_func
    }

    # Import and add any modules to the namespace
    for module_name in imports:
        module = __import__(module_name)
        namespace[module.__name__] = module

    # Exec the string to build the function and register it
    exec(func, namespace, namespace)

    return f"[{name}]"
