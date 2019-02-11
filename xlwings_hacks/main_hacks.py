import xlwings.main as xlmain
import sys

if sys.platform.startswith('win'):
    from . import _xlwindows_hacks as xlplatform_hacks
else:
    pass
    # from . import _xlmac_hacks as xlplatform_hacks
    # not yet implemented.


class BaseTable(object):
    def __init__(self, impl=None):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.
        """
        return self.impl.api

    @property
    def parent(self):
        """
        Returns the parent of the object.
        """
        return Sheet(impl=self.impl.parent)

    @property
    def name(self):
        """
        Returns or sets the name of the object.
        """
        return self.impl.name

    def refresh(self):
        self.impl.refresh()


class BaseTables(xlmain.Collection):
    @property
    def parent(self):
        return xlmain.Sheet(impl=self.impl.parent)

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.
        """
        return self.impl.api


class ListObject(BaseTable):
    @property
    def querytable(self):
        return QueryTable(self.impl.querytable)

    @property
    def showtotals(self):
        return self.impl.showtotals

    @showtotals.setter
    def showtotals(self, value):
        self.impl.showtotals = value

    @property
    def listcolumns(self):
        return ListColumns(self.impl.listcolumns)


class ListObjects(BaseTables):
    _wrap = ListObject


class BaseListRowColumn(object):
    """internal class."""
    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        return self.impl

    @property
    def parent(self):
        return ListObject(impl=self.impl.parent)


class BaseListRowsColumns(xlmain.Collection):
    """
    internal class.
    _wrap attribute must be not None.
    """
    def __init__(self, impl):
        xlmain.Collection.__init__(self, impl)

    @property
    def parent(self):
        return ListObjects(impl=self.impl.parent)


class ListColumn(BaseListRowColumn):
    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def totals_calculation(self):
        return self.impl.totals_calculation

    @totals_calculation.setter
    def totals_calculation(self, calculation):
        self.impl.totals_calculation = calculation


class ListColumns(BaseListRowsColumns):
    _wrap = ListColumn


class QueryTable(BaseTable):
    @property
    def background_query(self):
        return self.impl.background_query

    @background_query.setter
    def background_query(self, value):
        self.impl.background_query = value

    @property
    def command_text(self):
        return self.impl.command_text

    @command_text.setter
    def command_text(self, text):
        self.impl.command_text = text

    @property
    def listobject(self):
        return ListObject(self.impl.listobject)


class QueryTables(BaseTables):
    """
    A collection of all :meth:`querytable <QueryTable>` objects:

    Examples
    --------

    .. code-block:: python

        import xlwings as xw

        wb = xw.Book()
        ws = wb.sheets.add()
        rng = ws.range((1, 1))
        qts = QueryTables(xlplatform_hacks._attr_querytables(ws.impl))
    """
    _wrap = QueryTable

    def add(self, connection, destination, sql=None):
        """
        Creates a new QueryTable.

        Parameters
        ----------
        connection : str, ADO/DAO recordset, web query, data finder, text file
            A datasource of the table.
        destination : Range
            A range in the upper-left corner of the Sheet.
        sql : str, default None
            A SQL query str.

        -------
        """
        impl = self.impl.add(
            connection,
            destination.impl,
            sql)
        return self._wrap(impl)
