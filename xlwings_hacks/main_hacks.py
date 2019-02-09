import xlwings as xw
import xlwings.main as xlmain
import sys

if sys.platform.startswith('win'):
    import _xlwindows_hacks as xlplatform_hacks
else:
    pass
    # import _xlmac_hacks as xlplatform_hacks


class ListObject(object):
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
        Returns the parent of the listobject.
        """
        return Sheet(impl=self.impl.parent)

    @property
    def name(self):
        """
        Returns or sets the name of the listobject.
        """
        return self.impl.name


class ListObjects(xlmain.Collection):
    _wrap = ListObject

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


class SheetWithListObject(xw.Sheet):
    @property
    def listobjects(self):
        return ListObjects(
            xlplatform_hacks._sheet_attr_listobjects(
                self.impl))


if __name__ == '__main__':
    import xlwings as xw

    wb = xw.Book()
    ws = SheetWithListObject(impl=wb.sheets[0].impl)
    rng = ws.range((1, 1))

    # impl_los = xlplatform_hacks.ListObjects(ws.impl.xl.ListObjects)
    # los = ListObjects(impl_los)
    los = ws.listobjects

    print(len(los))

    myListObject = ws.api.ListObjects.Add(
        SourceType=0,
        Source="ODBC;"
        "DSN=MS Access Database;"
        "DBQ=C:\python\Access2003.mdb;"
        "DefaultDir=C:\python\Access2003.mdb;"
        "DriverId=25;"
        "FIL=MS Access;"
        "MaxBufferSize=2048;"
        "PageTimeout=5;",
        LinkSource=True,
        Destination=rng.api)
    myQueryTable = myListObject.QueryTable

    myQueryTable.CommandText = "SELECT * FROM 顧客名"

    myQueryTable.Refresh()

    for lo in los:
        print(lo.name)
