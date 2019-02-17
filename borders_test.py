import xlwings as xw
from xlwings_hacks.main_hacks import get_borders_of
from xlwings.constants import *


if __name__ == '__main__':
    wb = xw.Book()

    ws = wb.sheets[0]

    rng = ws.range((3, 3))
    bds = get_borders_of(rng)
    bd = bds["top"]

    bd.style = "continuous"
    bd.weight = "thick"

    bds = get_borders_of(ws.range((5, 2)))
    for bd in list(bds)[0:4]:
        bd.color = RgbColor.rgbRed
