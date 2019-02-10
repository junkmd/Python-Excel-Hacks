import xlwings.main as xlmain
import sys

if sys.platform.startswith('win'):
    import _xlwindows_hacks as xlplatform_hacks
else:
    pass
    # import _xlmac_hacks as xlplatform_hacks


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


class ListObjects(BaseTables):
    _wrap = ListObject


class SheetWithListObject(xlmain.Sheet):
    @property
    def listobjects(self):
        """
        Returns a ListObjects.
        """
        return ListObjects(
            xlplatform_hacks._attr_listobjects(
                self.impl))


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


class QueryTables(BaseTables):
    _wrap = QueryTable


if __name__ == '__main__':
    import xlwings as xw

    CONN_STR = \
        "ODBC;"\
        "DSN=MS Access Database;"\
        "DBQ=C:\python\Access2003.mdb;"\
        "DefaultDir=C:\python\Access2003.mdb;"\
        "DriverId=25;"\
        "FIL=MS Access;"\
        "MaxBufferSize=2048;"\
        "PageTimeout=5;"

    sqlstr = "SELECT * FROM 顧客名"

    wb = xw.Book()

    ws = SheetWithListObject(impl=wb.sheets.add().impl)
    los = ws.listobjects
    rng = ws.range((1, 1))

    myListObject = ws.api.ListObjects.Add(
        SourceType=0,
        Source=CONN_STR,
        LinkSource=True,
        Destination=rng.api)

    lo = los[0]
    lo.querytable.command_text = sqlstr
    lo.refresh()

    ws = wb.sheets.add()
    qts = QueryTables(xlplatform_hacks._attr_querytables(ws.impl))
    rng = ws.range((1, 1))

    myQueryTable = ws.api.QueryTables.Add(
        CONN_STR, rng.api, sqlstr)

    qt = qts[0]
    qt.refresh()
