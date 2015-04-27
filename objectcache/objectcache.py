"""
Object Cache Example
====================

Excel cells hold basic types (strings, numbers, booleans etc) but sometimes
it can be useful to have functions that take and return objects and to be
able to call those functions from Excel.

This example shows how a custom type ('cached_object') and an object cache
can be used to pass objects between functions using PyXLL.

It also shows how COM events can be used to remove items from
the object cache when they are no longer needed.
"""
from pyxll import (
    xlfCaller,
    xl_arg_type,
    xl_return_type,
    xl_func,
    xl_on_open,
    xl_on_close,
    xl_on_reload
)
import pyxll
import logging

_log = logging.getLogger(__name__)

# win32com and automation.xl_app are required for the code that
# cleans up the cache in response to Excel events.
# The basic ObjectCache and related code will work without these
# modules.
try:
    import win32com.client
    _have_win32com = True
except ImportError:
    _log.warning("win32com.client could not be imported")
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

        # set of open workbooks and opening workbooks
        self.__workbooks = set()
        self.__pending_workbooks = set()

    def __len__(self):
        """returns the number of cached objects"""
        return len(self.__objects)

    @staticmethod
    def _get_obj_id(obj):
        """returns the id for an object stored in the cache"""
        # the object id must be unique for objects within the cache
        cls_name = getattr(obj, "__class__", type(obj)).__name__
        return "<%s instance at 0x%x>" % (cls_name, id(obj))

    def add_workbook(self, workbook):
        """
        Called when a Workbook opens.
        Clears any cached objects from any previous workbooks with the same name.

        :param workbook: Workbook instance.
        """
        workbook_name = str(workbook.Name)

        if workbook_name not in self.__pending_workbooks:
            self.delete_all(workbook_name)

        self.__workbooks.add(workbook_name)
        self.__pending_workbooks.discard(workbook_name)

        # Would like to delete invalid cells since we may have closed/reopened
        # but need access to individual sheet objects
        self.delete_invalid(workbook)

    def update(self, workbook_name, sheet_name, cell, value):
        """
        Updates the cached value for a workbook, sheet and cell and returns the cache id.

        :param str workbook_name: Workbook name.
        :param str sheet_name: Worksheet name.
        :param cell: (row, col) tuple.
        :param value: Value to be cached.
        """
        obj_id = self._get_obj_id(value)

        # remove any previous entry in the cache for this cell
        self.delete(workbook_name, sheet_name, cell)

        _log.debug("Adding entry %s to cache at (%s, %s, %s)" % (
            obj_id, workbook_name, sheet_name, cell))

        # update the object cache to include this cell as a referring cell
        # (a dict is used instead of a set to be compatible with older python versions)
        unused, referring_cells = self.__objects.setdefault(obj_id, (value, {}))
        referring_cells[(workbook_name, sheet_name, cell)] = None

        # update the cache of cells to object ids
        self.__cells.setdefault(workbook_name, {}).setdefault(sheet_name, {})[cell] = obj_id

        # note that this workbook is opening if it's not in our set of workbooks
        if workbook_name not in self.__workbooks:
            self.__pending_workbooks.add(workbook_name)

        # return the id for fetching the object from the cache later
        return obj_id

    def get(self, obj_id):
        """
        Return an object stored in the cache by the object id returned
        from the update method.

        :param obj_id: Identifier returned by `update`.
        """
        try:
            return self.__objects[obj_id][0]
        except KeyError:
            raise ObjectCacheKeyError(obj_id)

    def delete(self, workbook_name, sheet_name, cell):
        """
        Deletes the cached value for a workbook, sheet and cell.

        :param str workbook_name: Workbook name.
        :param str sheet_name: Worksheet name.
        :param cell: (row, col) tuple.
        """
        try:
            obj_id = self.__cells[workbook_name][sheet_name][cell]
        except KeyError:
            # Nothing cached for this cell.
            return

        _log.debug("Removing entry %s from cache at (%s, %s, %s)" % (
            obj_id, workbook_name, sheet_name, cell))

        # Remove this cell from the object's referring cells and remove the
        # object from the cache if no more cells are referring to it.
        obj, referring_cells = self.__objects[obj_id]
        del referring_cells[(workbook_name, sheet_name, cell)]
        if not referring_cells:
            del self.__objects[obj_id]

        # Remove the entries from the __cells dict.
        wb_cache = self.__cells[workbook_name]
        ws_cache = wb_cache[sheet_name]
        del ws_cache[cell]
        if not ws_cache:
            del wb_cache[sheet_name]
        if not wb_cache:
            del self.__cells[workbook_name]

    def delete_all(self, workbook_name, sheet_name=None):
        """
        Delete all references in the cache by workbook, worksheet.

        :param str workbook_name: Workbook name.
        :param str sheet_name: Worksheet name.
        """
        wb_cache = self.__cells.get(workbook_name)
        if wb_cache is not None:
            if sheet_name is not None:
                sheet_names = [sheet_name]
            else:
                sheet_names = list(wb_cache.keys())

            for sheet_name in sheet_names:
                ws_cache = wb_cache.get(sheet_name)
                if ws_cache is not None:
                    for cell in list(ws_cache.keys()):
                        self.delete(workbook_name, sheet_name, cell)

    def delete_invalid(self, workbook, sheet=None):
        """
        Deletes all invalid references in the cache by workbook, worksheet.

        References are considered invalid when the cell contents no longer
        matches the cached object identifier of the cache object for that
        cell.

        Cached values can become invalid if a sheet or workbook is deleted,
        or if the contents of a cell that contained an object reference is
        overwritten.
        """
        workbook_name = workbook.Name
        wb_cache = self.__cells.get(workbook_name)
        if wb_cache is None:
            return

        if sheet is not None:
            sheet_names = [sheet.Name]
        else:
            sheet_names = wb_cache.keys()

        for sheet_name in sheet_names:
            ws_cache = wb_cache.get(sheet_name)
            if ws_cache is None:
                continue

            try:
                if sheet is None:
                    sheet = workbook.Worksheets(sheet_name)
                check_cell_contents = True
            except:
                check_cell_contents = False

            for cell, obj_id in list(ws_cache.items()):
                if check_cell_contents:
                    # The cell tuple is zero offset, but Cells expects them starting from one.
                    row, col = cell
                    cell_value = sheet.Cells(row+1, col+1).Value
                    if cell_value == obj_id:
                        continue

                # Either the sheet doesn't exist or the cell value
                # doesn't match the object id, so delete it from
                # the cache.
                self.delete(workbook_name, sheet_name, cell)


#
# There's one global instance of the cache.
#
_global_cache = ObjectCache()


#
# Here we register the functions that convert the cached objects to and
# from more basic types so they can be used by PyXLL Excel functions.
#

@xl_return_type("cached_object", "string", allow_arrays=False, thread_safe=False)
def cached_object_return_func(x):
    """
    Custom return type for objects that should be cached for use as
    parameters to other xl functions.
    """
    global _global_cache
    # This requires the function to be registered as a macro sheet equivalent
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
        # ensure the cache is kept up to date with cell changes.
        _setup_event_handler(_global_cache)

    # Get the calling sheet and cell.
    caller = xlfCaller()
    sheet = caller.sheet_name
    cell = (caller.rect.first_row, caller.rect.first_col)

    # Check the function isn't being used as an array function.
    assert cell == (caller.rect.last_row, caller.rect.last_col), \
        "Functions returning objects should not be used as array functions"

    # The sheet name will be in "[book]sheet" format.
    workbook = None
    if sheet.startswith("[") and "]" in sheet:
        workbook, sheet = sheet.strip("[").split("]", 1)

    # Update the cache and return the cached object id.
    return _global_cache.update(workbook, sheet, cell, x)


@xl_arg_type("cached_object", "string")
def cached_object_arg_func(x, thread_safe=False):
    """
    Custom argument type for objects that have been stored in the
    global object cache.
    """
    # Lookup the object in the cache by its cached object id.
    global _global_cache
    return _global_cache.get(x)


#
# Utility function to check how many objects are in the cache.
#
@xl_func(": int", volatile=True)
def cached_object_count():
    """Return the number of cached objects"""
    global _global_cache
    return len(_global_cache)


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
    @staticmethod
    def null_event_handler(*args, **kwargs):
        return None

    def __new__(mcs, name, bases, dict):
        # Construct the new class.
        cls = type.__new__(mcs, name, bases, dict)

        # Create dummy methods for any missing event handlers.
        cls._dispid_to_func_ = getattr(cls, "_dispid_to_func_", {})
        for dispid, name in cls._dispid_to_func_.iteritems():
            func = getattr(cls, name, None)
            if func is None:
                setattr(cls, name, EventHandlerMetaClass.null_event_handler)
        return cls


class ObjectCacheApplicationEventHandler(object):
    """
    An event handler for Application events used to clean entries from
    the object cache that would otherwise be missed.
    """
    __metaclass__ = EventHandlerMetaClass

    def __init__(self):
        # We have an event handler per workbook, but they only get
        # created once set_cache is called.
        self.__wb_event_handlers = {}
        self.__cache = None

    def set_cache(self, cache):
        self.__cache = cache

        # Create event handlers for all of the current workbooks.
        for workbook in self.Workbooks:
            wb = win32com.client.DispatchWithEvents(workbook, ObjectCacheWorkbookEventHandler)
            wb.set_cache(cache)
            self.__wb_event_handlers[workbook.Name] = wb

    def OnWorkbookOpen(self, workbook):
        # This workbook can't have anything in the cache yet, so make
        # sure it doesn't (it's possible a workbook with the same name
        # was closed with some cached entries and this one was then
        # opened).
        if self.__cache is not None:
            self.__cache.add_workbook(workbook)

            # Create a new workbook event handler for this workbook.
            wb = win32com.client.DispatchWithEvents(workbook, ObjectCacheWorkbookEventHandler)
            wb.set_cache(self.__cache)

            # Delete any previous handler now rather than possibly wait for the GC.
            if workbook.Name in list(self.__wb_event_handlers):
                del self.__wb_event_handlers[workbook.Name]
                self.__wb_event_handlers[workbook.Name] = wb

    def OnWorkbookActivate(self, workbook):
        # Remove any workbooks that no longer exist.
        wb_names = [x.Name for x in self.Workbooks]
        for name in list(self.__wb_event_handlers.keys()):
            if name not in wb_names:
                # It's gone so remove the cache entries and the wb handler.
                if self.__cache is not None:
                    self.__cache.delete_all(str(name))
                del self.__wb_event_handlers[name]

        # Add in any new workbooks, which can happen if a workbook has just been renamed.
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
        # Keep track of sheets we know about for when sheets get deleted or renamed.
        self.__sheets = [x.Name for x in self.Sheets]
        self.__cache = None

    def set_cache(self, cache):
        self.__cache = cache

    def OnWorkbookNewSheet(self, sheet):
        # This work can't have anything in the cache yet.
        if self.__cache is not None:
            self.__cache.delete_all(str(self.Name), str(sheet.Name))

        # Add it to our list of known sheets.
        self.__sheets.append(sheet.Name)

    def OnSheetActivate(self, sheet):
        # Remove any worksheets that not longer exist.
        ws_names = [x.Name for x in self.Sheets]
        for name in list(self.__sheets):
            if name not in ws_names:
                # It's gone so remove the cache entries an the reference.
                if self.__cache is not None:
                    self.__cache.delete_all(str(self.Name), str(name))
                self.__sheets.remove(name)

        # Ensure our list includes any new names due to renames.
        self.__sheets = ws_names

    def OnSheetChange(self, sheet, change_range):
        if self.__cache is not None:
            self.__cache.delete_invalid(self, sheet)


def _xl_app():
    """Return a Dispatch object for the current Excel instance."""
    # Get the Excel application object from PyXLL and wrap it.
    xl_window = pyxll.get_active_object()
    xl_app = win32com.client.Dispatch(xl_window).Application

    # It's helpful to make sure the gen_py wrapper has been created
    # as otherwise things like constants and event handlers won't work.
    win32com.client.gencache.EnsureDispatch(xl_app)

    return xl_app


_event_handlers = {}
def _setup_event_handler(cache):
    # Only setup the app event handler once.
    if cache not in _event_handlers:
        app_handler = win32com.client.DispatchWithEvents(_xl_app(),
                                                         ObjectCacheApplicationEventHandler)
        app_handler.set_cache(cache)
        _event_handlers[cache] = app_handler


@xl_on_open
def _startup(*args):
    _setup_event_handler(_global_cache)


@xl_on_reload
@xl_on_close
def _delete_event_handlers(*args):
    # Make sure the event handles are deleted now as otherwise they could still
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
