'''xlwings hacks'''
import xlwings._xlwindows as xlwindows
from abc import ABC, ABCMeta, abstractmethod


class BaseTable(object):
    """internal class."""
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def api(self):
        return self.xl

    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def parent(self):
        return Sheet(xl=self.xl.Parent)

    def refresh(self):
        self.xl.Refresh()


class BaseTables(xlwindows.Collection):
    """
    internal class.
    _wrap attribute must be not None.
    """
    def __init__(self, xl):
        xlwindows.Collection.__init__(self, xl)

    @property
    def parent(self):
        return xlwindows.Sheet(xl=self.xl.Parent)


def _attr_tables(py_class, xl_obj):
    """
    internal function.
    Returns the tables implement of the object.

    Arguments
    ---------
    py_class : Class of implement.
    xl_obj : win32com object.

    Examples
    --------

    .. code-block:: python

        def _attr_listobjects(obj):
            return _attr_tables(ListObjects, obj.xl.ListObjects)
    """
    return py_class(xl_obj)


class ListObject(BaseTable):
    @property
    def querytable(self):
        return QueryTable(self.xl.QueryTable)


class ListObjects(BaseTables):
    _wrap = ListObject


def _attr_listobjects(obj):
    return _attr_tables(ListObjects, obj.xl.ListObjects)


class QueryTable(BaseTable):
    def __init__(self, xl):
        BaseTable.__init__(self, xl)
        self.background_query = False

    @property
    def background_query(self):
        return self.xl.BackgroundQuery

    @background_query.setter
    def background_query(self, value):
        self.xl.BackgroundQuery = value

    @property
    def command_text(self):
        return self.xl.CommandText

    @command_text.setter
    def command_text(self, text):
        self.xl.CommandText = text


class QueryTables(BaseTables):
    _wrap = QueryTable


def _attr_querytables(obj):
    return _attr_tables(QueryTables, obj.xl.QueryTables)


if __name__ == '__main__':
    import xlwings as xw

    pass
