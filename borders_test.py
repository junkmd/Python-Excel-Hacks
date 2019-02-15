import xlwings as xw
from xlwings_hacks.main_hacks import get_borders_of
from xlwings.constants import *


if __name__ == '__main__':
    wb = xw.Book()

    ws = wb.sheets[0]

    rng = ws.range((2, 2))

    bds = get_borders_of(rng)

    bds_ids = [
        BordersIndex.xlDiagonalDown,
        BordersIndex.xlDiagonalUp,
        BordersIndex.xlEdgeBottom,
        BordersIndex.xlEdgeLeft,
        BordersIndex.xlEdgeRight,
        BordersIndex.xlEdgeTop,
        BordersIndex.xlInsideHorizontal,
        BordersIndex.xlInsideVertical]

    styles = [
        LineStyle.xlContinuous,
        LineStyle.xlDash,
        LineStyle.xlDashDot,
        LineStyle.xlDashDotDot,
        LineStyle.xlDot,
        LineStyle.xlSlantDashDot
    ]

    i = 0

    for bd, style in zip(bds, styles):
        bd.api.LineStyle = style
        # bd.api.Weight = BorderWeight.xlThin
        bd.weight = "thick"
        print(i)
        i += 1

    rng = ws.range((3, 3))
    bds = get_borders_of(rng)
    bd = bds["top"]

    bd.style = "continuous"
    bd.weight = "thick"
