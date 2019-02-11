import xlwings.main as xlmain
import sys

if sys.platform.startswith('win'):
    from . import _xlwindows_hacks as xlplatform_hacks
else:
    pass
    # from . import _xlmac_hacks as xlplatform_hacks
    # not yet implemented.


class BaseTable(object):
    """internal class."""
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
        """
        Updates an external data range.
        For only based on the results of a SQL query.
        """
        self.impl.refresh()


class BaseTables(xlmain.Collection):
    """internal class."""
    @property
    def parent(self):
        """
        Returns the parent of the object.
        """
        return xlmain.Sheet(impl=self.impl.parent)

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.
        """
        return self.impl.api


class ListObject(BaseTable):
    """
    Represents a ListObject object.
    An Object in the ListObjects collection.
    """
    @property
    def querytable(self):
        """
        Returns the QueryTable object that provides a link
        for the ListObject object to the list server.
        """
        return QueryTable(self.impl.querytable)

    @property
    def showtotals(self):
        """
        Gets or sets whether the Total row is visible.
        """
        return self.impl.showtotals

    @showtotals.setter
    def showtotals(self, value):
        self.impl.showtotals = value

    @property
    def listcolumns(self):
        """
        Returns a ListColumns collection that represents
        all the columns in a ListObject object.
        """
        return ListColumns(self.impl.listcolumns)

    def unlink(self):
        """
        Removes the link to a DB et al.
        """
        self.impl.unlink()


class ListObjects(BaseTables):
    """
    A collection of all the ListObject objects on a worksheet.
    Each ListObject object represents a table in the worksheet.
    """
    _wrap = ListObject

    def add(self, source_type, source, destination=None, has_headers='guess'):
        """
        Creates a new list object.

        Parameters
        ----------
        source_type : 'external', 'query', 'range' or 'xml'
            Indicates the kind of source for the query.
        source : str or Range.
            If source_type was 'range', must be Range.
        destination : Range or None.
            If source_type was 'range', must be None.
        has_headers : 'yes', 'no' or default 'guess'
            'guess' : Excel determines whether there is a header.
            'yes' : Top row of range will be header.
            'no' : The header row will be added to top of the entire range.

        -------
        """
        if destination is not None:
            dest = destination.impl
        else:
            dest = None

        if source_type == 'range':
            src = source.impl
        else:
            src = source

        return ListObject(
            self.impl.add(
                source_type, src, dest, has_headers)
            )


class BaseListRowColumn(object):
    """internal class."""
    def __init__(self, impl):
        self.impl = impl

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.
        """
        return self.impl

    @property
    def parent(self):
        """
        Returns the parent of the object.
        """
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
        """
        Returns the parent of the object.
        """
        return ListObjects(impl=self.impl.parent)


class ListColumn(BaseListRowColumn):
    """
    Represents a column in a table.
    """
    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def totals_calculation(self):
        """
        Gets or sets determineing the type of calculation in the Totals row.
        """
        return self.impl.totals_calculation

    @totals_calculation.setter
    def totals_calculation(self, calculation):
        self.impl.totals_calculation = calculation


class ListColumns(BaseListRowsColumns):
    """
    A collection of all the ListColumn objects in the specified ListObject.
    """
    _wrap = ListColumn


class QueryTable(BaseTable):
    """
    Represents a QueryTable object.
    An Object in the QueryTables collection.
    """
    @property
    def background_query(self):
        """
        Gets or sets the performance of refreshing to True or False.

            True: The query table are performed asynchronously.
            False: The query table are NOT performed asynchronously.
        """
        return self.impl.background_query

    @background_query.setter
    def background_query(self, value):
        self.impl.background_query = value

    @property
    def command_text(self):
        """
        Returns or sets the command string for the data source.
        """
        return self.impl.command_text

    @command_text.setter
    def command_text(self, text):
        self.impl.command_text = text

    @property
    def listobject(self):
        """
        Returns a ListObject object for the QueryTable object.
        """
        return ListObject(self.impl.listobject)


class QueryTables(BaseTables):
    """
    A collection of all QueryTable objects on a worksheet.
    Each QueryTable object represents a table in the worksheet.

    Examples
    --------

    .. code-block:: python

        import xlwings as xw

        # now rewriting...
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


# --- implemented sheet ---
class Sheet_Hacked(xlmain.Sheet):
    """
    Hacked xlwings.main.Sheet
    """
    def __init__(self, impl):
        xlmain.Sheet.__init__(self, impl=impl)

    @property
    def listobjects(self):
        """
        A collection of all the ListObject objects on a worksheet.
        Each ListObject object represents a table in the worksheet.
        """
        return ListObjects(
            impl=xlplatform_hacks._attr_listobjects(self.impl)
        )

    @property
    def querytables(self):
        """
        Represents a worksheet table built from data returned from
        an external data source,
        such as an SQL server or a Microsoft Access database.
        """
        return QueryTables(
            impl=xlplatform_hacks._attr_querytables(self.impl)
        )
