'''xlwings hacks'''
import xlwings._xlwindows as xlwindows


class ListObject(object):
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


class ListObjects(xlwindows.Collection):

    _wrap = ListObject

    def __init__(self, xl):
        xlwindows.Collection.__init__(self, xl)

    @property
    def parent(self):
        return xlwindows.Sheet(xl=self.xl.Parent)

if __name__ == '__main__':
    import xlwings as xw

    wb = xw.Book()
    ws = wb.sheets[0]
    rng = ws.range((1, 1))

    impl_los = ListObjects(ws.impl.xl.ListObjects)

    print(len(impl_los))

    myListObject = ws.impl.xl.ListObjects.Add(
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

    for impl_lo in impl_los:
        print(impl_lo.name)
