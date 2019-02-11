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

    @property
    def showtotals(self):
        return self.xl.ShowTotals

    @showtotals.setter
    def showtotals(self, value):
        self.xl.ShowTotals = value

    @property
    def listcolumns(self):
        return ListColumns(self.xl.ListColumns)


class ListObjects(BaseTables):
    _wrap = ListObject


def _attr_listobjects(obj):
    return _attr_tables(ListObjects, obj.xl.ListObjects)


class BaseListRowColumn(object):
    """internal class."""
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return ListObject(xl=self.xl.Parent)


class BaseListRowsColumns(xlwindows.Collection):
    """
    internal class.
    _wrap attribute must be not None.
    """
    def __init__(self, xl):
        xlwindows.Collection.__init__(self, xl)

    @property
    def parent(self):
        return ListObjects(xl=self.xl.Parent)


class ListColumn(BaseListRowColumn):
    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def totals_calculation(self):
        return totals_i2s[self.xl.TotalsCalculation]

    @totals_calculation.setter
    def totals_calculation(self, calculation):
        self.xl.TotalsCalculation = totals_s2i[calculation]


class ListColumns(BaseListRowsColumns):
    _wrap = ListColumn


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

    @property
    def listobject(self):
        return ListObject(self.xl.ListObject)


class QueryTables(BaseTables):
    _wrap = QueryTable

    def add(self, connection, destination, sql=None):
        if sql is None:
            xl = self.xl.Add(
                Connection=connection,
                Destination=destination.api)
        else:
            xl = self.xl.Add(
                Connection=connection,
                Destination=destination.api,
                Sql=sql)
        return QueryTable(xl)


def _attr_querytables(obj):
    return _attr_tables(QueryTables, obj.xl.QueryTables)

# --- constants ---
totals_s2i = {
    'average': 2,  # TotalsCalculation.xlTotalsCalculationAverage
    'count': 3,  # TotalsCalculation.xlTotalsCalculationCount
    'countnums': 4,  # TotalsCalculation.xlTotalsCalculationCountNums
    'custom': 9,  # TotalsCalculation.xlTotalsCalculationCustom
    'max': 6,  # TotalsCalculation.xlTotalsCalculationMax
    'min': 5,  # TotalsCalculation.xlTotalsCalculationMin
    'none': 0,  # TotalsCalculation.xlTotalsCalculationNone
    'stddev': 7,  # TotalsCalculation.xlTotalsCalculationStdDev
    'sum': 1,  # TotalsCalculation.xlTotalsCalculationSum
    'var': 8  # TotalsCalculation.xlTotalsCalculationVar
    }

totals_i2s = {v: k for k, v in totals_s2i.items()}
