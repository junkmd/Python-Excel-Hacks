import xlwings as xw
from xlwings_hacks.main_hacks import Sheet_Hacked, ListObject
from xlwings.constants import *
from pprint import pformat


if __name__ == '__main__':

    MDB_FILE_FULLPATH = r"C:\python\Access2003.mdb"

    ACCESS_CONN_STR = \
        "ODBC;"\
        "DSN=MS Access Database;"\
        "DBQ=%s;"\
        "DefaultDir=%s;"\
        "DriverId=25;"\
        "FIL=MS Access;"\
        "MaxBufferSize=2048;"\
        "PageTimeout=5;"

    CONN_STR = ACCESS_CONN_STR % (MDB_FILE_FULLPATH, MDB_FILE_FULLPATH)

    sqlstr = "SELECT * FROM Test_Table"

    wb = xw.Book()

    ws = Sheet_Hacked(impl=wb.sheets.add().impl)
    los = ws.listobjects
    rng = ws.range((1, 1))

    lo = los.add('external', CONN_STR, rng)
    lo.querytable.command_text = sqlstr
    lo.refresh()
    lo.querytable.listobject.showtotals = True

    for lc in lo.listcolumns:
        lc.name = lc.name + "_"
        lc.totals_calculation = "none"
        lc.range.number_format = "0!.0,\"万\";[赤]-0!.0,\"万\";0"

    ws = Sheet_Hacked(impl=wb.sheets.add().impl)
    qts = ws.querytables
    rng = ws.range((1, 1))

    qt = qts.add(CONN_STR, rng, sqlstr)

    qt.refresh()

    ws = Sheet_Hacked(impl=wb.sheets.add().impl)
    for i in range(1, 4):
        ws.range(1, i).value = "header_%s" % i
    lo = ws.listobjects.add(
        'range', ws.range((1, 1), (3, 3)), has_headers='yes')

    lr = lo.listrows.add(0)

    for clr in lr.columns:
        print(clr.column_name, clr)

    with ws.pagesetup as psu:
        psu.fit_to_tall = 1
        psu.fit_to_wide = 0
