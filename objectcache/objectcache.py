"""
Excel cells hold basic types (strings, numbers, booleans etc) but sometimes
it can be useful to have functions that take and return objects and to be
able to call those functions from Excel.

This example shows how a custom type ('cached_object') and an object cache
can be used to pass objects between functions using PyXLL.

It also shows how COM events can be used to remove items from
the object cache when they are no longer needed.

For information about custom types in PyXLL see:
https://www.pyxll.com/docs/udfs.html#custom-types
"""

from pyxll import xlfCaller,        \
                    xl_arg_type,    \
                    xl_return_type, \
                    xl_func,        \
                    xl_on_close,    \
                    xl_on_reload
import pyxll

import logging
_log = logging.getLogger(__name__)

#
# win32com and automation.xl_app are required for the code that
# cleans up the cache in response to Excel events.
# The basic ObjectCache and related code will work without these
# modules.
try:
    import win32com.client
    _have_win32com = True
except ImportError:
    _log.warning("*** win32com.client could not be imported           ***")
    _log.warning("*** some of the objectcache examples will not work  ***")
    _log.warning("*** to fix this, install the pywin32 extensions     ***")
    _have_win32com = False

class ObjectCacheKeyError(KeyError):
    """
    Exception raised when attempting to retrieve an object from the
    cache that's not found.
    """
    def __init__(self, key):
        KeyError.__init__(self, key)

class ObjectCache(object):
    """
    ObjectCache maintains a cache of objects returned to Excel
    and the cells referring to those objects.
    
    As xl functions return objects they update the cache and
    any previously cached objects are removed from the cache
    when they are no longer referred to by any cells.
    
    Custom functions don't reference this class directly,
    instead they use the custom type 'cached_object' which
    is registered with PyXLL after this class.
    """

    def __init__(self):
        # dict of workbooks -> worksheets -> cell to object ids
        self.__cells = {}
        
        # dict of object ids to (object, {[referring (wb, ws, cell)] -> None})
        self.__objects = {}

    def __len__(self):
        """returns the number of cached objects"""
        return len(self.__objects)

    @staticmethod
    def _get_obj_id(obj):
        """returns the id for an object stored in the cache"""
        # the object id must be unique for objects within the cache
        cls_name = getattr(obj, "__class__", type(obj)).__name__
        return "<%s instance at 0x%x>" % (cls_name, id(obj))
        
    def update(self, workbook, sheet, cell, value):
        """updates the cached value for a workbook, sheet and cell and returns the cache id"""
        obj_id = self._get_obj_id(value)

        # remove any previous entry in the cache for this cell
        self.delete(workbook, sheet, cell)

        _log.debug("Adding entry %s to cache at (%s, %s, %s)" % (obj_id, workbook, sheet, cell))

        # update the object cache to include this cell as a referring cell
        # (a dict is used instead of a set to be compatible with older python versions)
        unused, referring_cells = self.__objects.setdefault(obj_id, (value, {}))
        referring_cells[(workbook, sheet, cell)] = None

        # update the cache of cells to object ids
        self.__cells.setdefault(workbook, {}).setdefault(sheet, {})[cell] = obj_id

        # return the id for fetching the object from the cache later
        return obj_id

    def get(self, obj_id):
        """
        returns an object stored in the cache by the object id returned
        from the update method.
        """
        try:
            return self.__objects[obj_id][0]
        except KeyError:
            raise ObjectCacheKeyError(obj_id)

    def delete(self, workbook, sheet, cell):
        """deletes the cached value for a workbook, sheet and cell"""
        try:
            obj_id = self.__cells[workbook][sheet][cell]
        except KeyError:
            # nothing cached for this cell
            return

        _log.debug("Removing entry %s from cache at (%s, %s, %s)" % (obj_id, workbook, sheet, cell))

        # remove this cell from the object's referring cells and remove the
        # object from the cache if no more cells are referring to it
        obj, referring_cells = self.__objects[obj_id]
        del referring_cells[(workbook, sheet, cell)]
        if not referring_cells:
            del self.__objects[obj_id]

        # remove the entries from the __cells dict
        wb_cache = self.__cells[workbook]
        ws_cache = wb_cache[sheet]
        del ws_cache[cell]
        if not ws_cache:
            del wb_cache[sheet]
        if not wb_cache:
            del self.__cells[workbook]

    def delete_all(self, workbook, sheet=None, predicate=None):
        """
        deletes all references in the cache by workbook, worksheet.
        If predicate is not None, the cells will only be deleted if
        predicate(cell, obj_id) returns True
        """
        wb_cache = self.__cells.get(workbook)
        if wb_cache is not None:
            if sheet is not None:
                sheets = [sheet]
            else:
                sheets = wb_cache.keys()

            for sheet in sheets:
                ws_cache = wb_cache.get(sheet)
                if ws_cache is not None:
                    cached_cells = ws_cache.items()
                    for cell, obj_id in cached_cells:
                        if predicate is None or predicate(cell, obj_id):
                            self.delete(workbook, sheet, cell)

#
# there's one global instance of the cache
#
_global_cache = ObjectCache()

#
# Here we register the functions that convert the cached objects to and
# from more basic types so they can be used by PyXLL Excel functions
#

@xl_return_type("cached_object", "string", macro=True, allow_arrays=False, thread_safe=False)
def cached_object_return_func(x):
    """
    custom return type for objects that should be cached for use as
    parameters to other xl functions
    """
    # this requires the function to be registered as a macro sheet equivalent
    # function because it calls xlfCaller, hence macro=True in
    # the xl_return_type decorator above.
    #
    # As xlfCaller returns the individual cell a function was called from, it's
    # not possible to return arrays of cached_objects using the cached_object[] 
    # type in a function signature. allow_arrays=False prevents a function from
    # being registered with that return type. Arrays of cached_objects as an
    # argument type is fine though.

    if _have_win32com:
        # _setup_event_handler creates an event handler for Excel events to
        # ensure the cache is kept up to date with cell changes
        _setup_event_handler(_global_cache)

    # get the calling cell in [book]sheet!address format
    caller = xlfCaller()
    address = caller.address

    # split the cell up into workbook, sheet and cell
    assert "!" in address, "Calling cell address not in [book]sheet!address format: %s" % address
    wb_and_sheet, cell = address.split("!", 1)
    wb_and_sheet = wb_and_sheet.strip("'")

    assert wb_and_sheet.startswith("[") and "]" in wb_and_sheet, \
        "Calling cell not in [book]sheet!address format: %s" % address
    workbook, sheet = wb_and_sheet.strip("[").split("]", 1)
    while "''" in sheet:
        sheet = sheet.replace("''", "'")

    # update the cache and return the cached object id
    return _global_cache.update(workbook, sheet, cell, x)

@xl_arg_type("cached_object", "string")
def cached_object_arg_func(x, thread_safe=False):
    """
    custom argument type for objects that have been stored in the
    global object cache.
    """
    # lookup the object in the cache by its cached object id
    return _global_cache.get(x)


#
# Example worksheet functions using the object cache
#
# The following examples show how worksheet functions using
# xl_func can use the new 'cached_object' type registered
# above to return and take python objects cached by the
# object cache (appear to be cached on the excel grid).
#

@xl_func(": int", volatile=True)
def cached_object_count():
    """returns the number of cached objects"""
    return len(_global_cache)

class MyTestClass(object):
    """A basic class for testing the cached_object type"""

    def __init__(self, x):
        self.__x = x
 
    def __str__(self):
        return "%s(%s)" % (self.__class__.__name__, self.__x)
 
@xl_func("var: cached_object")
def cached_object_return_test(x):
    """returns an instance of MyTestClass"""
    return MyTestClass(x)

@xl_func("cached_object: string")
def cached_object_arg_test(x):
    """takes a MyTestClass instance and returns a string"""
    return str(x)
 
class MyDataGrid(object):
    """
    A second class for demonstrating cached_object types.
    This class is constructed with a grid of data and has
    some basic methods which are also exposed as worksheet
    functions.
    """
    
    def __init__(self, grid):
        self.__grid = grid
        
    def sum(self):
        """returns the sum of the numbers in the grid"""
        total = 0
        for row in self.__grid:
            total += sum(row)
        return total
        
    def __len__(self):
        total = 0
        for row in self.__grid:
            total += len(row)
        return total
        
    def __str__(self):
        return "%s(%d values)" % (self.__class__.__name__, len(self))
 
@xl_func("float[]: cached_object")
def make_datagrid(x):
    """returns a MyDataGrid object"""
    return MyDataGrid(x)
    
@xl_func("cached_object: int")
def datagrid_len(x):
    """returns the length of a MyDataGrid object"""
    return len(x)

@xl_func("cached_object: float")
def datagrid_sum(x):
    """returns the sum of a MyDataGrid object"""
    return x.sum()
 
@xl_func("cached_object: string")
def datagrid_str(x):
    """returns the string representation of a MyDataGrid object"""
    return str(x)
 
#
# So far we can cache objects and keep the cache up to date as
# functions are called and the return values change.
#
# However, if a cell is changed from a function that returns a cached
# object to something that doesn't there will be a reference
# left in the cache - and so references can be leaked. Or, if a workbook
# or worksheet is deleted objects will be leaked.
#
# We can hook into some of Excel's Application and Workbook events to
# detect when references to objects are no longer required and remove
# them from the cache.
#

class EventHandlerMetaClass(type):
    """
    A meta class for event handlers that don't repsond to all events.
    Without this an error would be raised by win32com when it tries
    to call an event handler method that isn't defined by the event
    handler instance.
    """

    def __new__(mcs, name, bases, dict):
        # construct the new class
        cls = type.__new__(mcs, name, bases, dict)

        # create dummy methods for any missing event handlers
        cls._dispid_to_func_ = getattr(cls, "_dispid_to_func_", {})
        for dispid, name in cls._dispid_to_func_.iteritems():
            func = getattr(cls, name, None)
            if func is None:
                func = lambda *args, **kwargs: None
                setattr(cls, name, func)

        return cls

class ObjectCacheApplicationEventHandler(object):
    """
    An event handler for Application events used to clean entries from
    the object cache that would otherwise be missed.
    """
    __metaclass__ = EventHandlerMetaClass

    def __init__(self):
        # we have an event handler per workbook, but they only get
        # created once set_cache is called.
        self.__wb_event_handlers = {}
        self.__cache = None

    def set_cache(self, cache):
        self.__cache = cache

        # create event handlers for all of the current workbooks
        for workbook in self.Workbooks:
            wb = win32com.client.DispatchWithEvents(workbook, ObjectCacheWorkbookEventHandler)
            wb.set_cache(cache)
            self.__wb_event_handlers[workbook.Name] = wb

    def OnWorkbookOpen(self, workbook):
        # this workbook can't have anything in the cache yet, so make
        # sure it doesn't (it's possible a workbook with the same name
        # was closed with some cached entries and this one was then
        # opened)
        if self.__cache is not None:
            self.__cache.delete_all(workbook=str(workbook.Name))

            # create a new workbook event handler for this workbook
            wb = win32com.client.DispatchWithEvents(workbook, ObjectCacheWorkbookEventHandler)
            wb.set_cache(self.__cache)

            # delete any previous handler now rather than possibly wait for the GC
            if workbook.Name in self.__wb_event_handlers:
                del self.__wb_event_handlers[workbook.Name]

            self.__wb_event_handlers[workbook.Name] = wb

    def OnWorkbookActivate(self, workbook):
        # remove any workbooks that no longer exist
        wb_names = [x.Name for x in self.Workbooks]
        for name, handler in self.__wb_event_handlers.items():
            if name not in wb_names:
                # it's gone so remove the cache entries and the wb handler
                if self.__cache is not None:
                    self.__cache.delete_all(str(name))
                del self.__wb_event_handlers[name]

        # add in any new workbooks, which can happen if a workbook has just been renamed
        if self.__cache is not None:
            for wb in self.Workbooks:
                if wb.Name not in self.__wb_event_handlers:
                    wb = win32com.client.DispatchWithEvents(wb, ObjectCacheWorkbookEventHandler)
                    wb.set_cache(self.__cache)
                    self.__wb_event_handlers[wb.Name] = wb

class ObjectCacheWorkbookEventHandler(object):
    """
    An event handler for Workbook events used to clean entries from
    the object cache that would otherwise be missed.
    """
    __metaclass__ = EventHandlerMetaClass

    def __init__(self):
        # keep track of sheets we know about for when sheets get deleted or renamed
        self.__sheets = [x.Name for x in self.Sheets]
        self.__cache = None

    def set_cache(self, cache):
        self.__cache = cache

    def OnWorkbookNewSheet(self, sheet):
        # this work can't have anything in the cache yet
        if self.__cache is not None:
            self.__cache.delete_all(str(self.Name), str(sheet.Name))
            
        # add it to our list of known sheets
        self.__sheets.append(sheet.Name)

    def OnSheetActivate(self, sheet):
        # remove any worksheets that not longer exist
        ws_names = [x.Name for x in self.Sheets]
        for name in list(self.__sheets):
            if name not in ws_names:
                # it's gone so remove the cache entries an the reference
                if self.__cache is not None:
                    self.__cache.delete_all(str(self.Name), str(name))
                self.__sheets.remove(name)

        # ensure our list includes any new names due to renames
        self.__sheets = ws_names

    def OnSheetChange(self, sheet, range):
        # delete all the cells from the cache where the cell is in range
        # and the current value is not the cached object id
        def check_cell(cell, obj_id):
            # check this cell is in the range that's changed
            cell = sheet.Range(cell)
            if range.Find(cell) is None:
                return False
            # check the cell's value has changed from obj_id
            return str(cell.Value) != obj_id

        if self.__cache is not None:
            self.__cache.delete_all(str(self.Name), str(sheet.Name), predicate=check_cell)

def _xl_app():
    """returns a Dispatch object for the current Excel instance"""
    # get the Excel application object from PyXLL and wrap it
    xl_window = pyxll.get_active_object()
    xl_app = win32com.client.Dispatch(xl_window).Application

    # it's helpful to make sure the gen_py wrapper has been created
    # as otherwise things like constants and event handlers won't work.
    win32com.client.gencache.EnsureDispatch(xl_app)

    return xl_app

_event_handlers = {}
def _setup_event_handler(cache):
    # only setup the app event handler once
    if cache not in _event_handlers:
        app_handler = win32com.client.DispatchWithEvents(_xl_app(),
                                                         ObjectCacheApplicationEventHandler)
        app_handler.set_cache(cache)
        _event_handlers[cache] = app_handler

@xl_on_reload
@xl_on_close
def _delete_event_handlers(*args):
    # make sure the event handles are deleted now as otherwise they could still
    # exist for a while until the GC gets to them, which can stop Excel from closing
    # or result in old event handlers still running if this module is reloaded.
    #
    # If you never wanted to reload this module, you could just import it from another
    # module loaded by pyxll and remove it from the pyxll.cfg and remove the
    # @xl_on_reload callback.
    #
    global _event_handlers
    handlers = _event_handlers.values()
    _event_handlers = {}
    while handlers:
        handler = handlers.pop()
        del handler
