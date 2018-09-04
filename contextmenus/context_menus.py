"""
Callbacks for context menus example.
Requires PyXLL >= 4 and Excel >= 2010.
See ribbon.xml.
"""
from pyxll import xl_app

def toggle_case(control):
    """Toggle the case of the currently selected cells"""
    # get the Excel Application object
    xl = xl_app()

    # iterate over the currently selected cells
    for cell in xl.Selection:
        # get the cell value
        value = cell.Value

        # skip any cells that don't contain text
        if not isinstance(value, str):
            continue

        # toggle between upper, lower and proper case
        if value.isupper():
            value = value.lower()
        elif value.islower():
            value = value.title()
        else:
            value = value.upper()

        # set the modified value on the cell
        cell.Value = value


def tolower(control):
    """Set the currently selected cells to lower case"""
    # get the Excel Application object
    xl = xl_app()

    # iterate over the currently selected cells
    for cell in xl.Selection:
        # get the cell value
        value = cell.Value

        # skip any cells that don't contain text
        if not isinstance(value, str):
            continue

        cell.Value = value.lower()


def toupper(control):
    """Set the currently selected cells to upper case"""
    # get the Excel Application object
    xl = xl_app()

    # iterate over the currently selected cells
    for cell in xl.Selection:
        # get the cell value
        value = cell.Value

        # skip any cells that don't contain text
        if not isinstance(value, str):
            continue

        cell.Value = value.upper()


def toproper(control):
    """Set the currently selected cells to 'proper' case"""
    # get the Excel Application object
    xl = xl_app()

    # iterate over the currently selected cells
    for cell in xl.Selection:
        # get the cell value
        value = cell.Value

        # skip any cells that don't contain text
        if not isinstance(value, str):
            continue

        cell.Value = value.title()


def dynamic_menu(control):
    """Return an xml fragment for the dynamic menu"""
    xml = """
         <menu xmlns="http://schemas.microsoft.com/office/2009/07/customui">
            <button id="Menu2Button1" label="Upper Case"
                imageMso="U"
                onAction="context_menus.toupper"/>
 
            <button id="Menu2Button2" label="Lower Case"
                imageMso="L"
                onAction="context_menus.tolower"/>
 
            <button id="Menu2Button3" label="Proper Case"
                imageMso="P"
                onAction="context_menus.toproper"/>
         </menu>
    """
    return xml
