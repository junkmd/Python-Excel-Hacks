
if __name__ == '__main__':
    import xlwings as xw
    from xlwings_hacks.main_hacks import xlplatform_hacks, \
        ListObjects, QueryTables

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

    ws = xw.Sheet(impl=wb.sheets.add().impl)
    los = ListObjects(xlplatform_hacks._attr_listobjects(ws.impl))
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

    qt = qts.add(CONN_STR, rng, sqlstr)

    qt.refresh()
