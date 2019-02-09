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


def _sheet_attr_listobjects(sheet):
    """
    Returns the listobjects implement of the sheet.

    Arguments
    ---------
    sheet : xlwings.Sheet.impl or xlwings.xlplatform.Sheet

    Examples
    --------

    .. code-block:: python

        import xlwings as xw

        class SubclassOfSheet(xw.Sheet):
            @property
            def listobjects(self):
                return _sheet_attr_listobjects(self.impl)

        wb = xw.Book()
        ws = SubclassOfSheet(impl=wb.sheets[0].impl)
        lo = ws.listobjects

    """
    return ListObjects(sheet.xl.ListObjects)
