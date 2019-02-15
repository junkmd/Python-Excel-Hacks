'''xlwings hacks'''
import xlwings._xlwindows as xlwindows
from abc import ABC, ABCMeta, abstractmethod


# --- Base of ListObject and QueryTable ---
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
        return xlwindows.Sheet(xl=self.xl.Parent)

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


def _attr_object_impl(py_class, xl_obj):
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
            return _attr_object_impl(ListObjects, obj.xl.ListObjects)
    """
    return py_class(xl_obj)


# --- ListObject ---
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

    def unlink():
        self.xl.Unlink()

    @property
    def range(self):
        return xlwindows.Range(xl=self.xl.Range)

    @property
    def header_row(self):
        return xlwindows.Range(xl=self.xl.HeaderRowRange)

    @property
    def body(self):
        return xlwindows.Range(xl=self.xl.DataBodyRange)

    @property
    def totals_row(self):
        return xlwindows.Range(xl=self.xl.TotalsRowRange)


class ListObjects(BaseTables):
    _wrap = ListObject

    def add(self, source_type, source, destination=None, has_headers='guess'):
        if source_type == 'range':
            xl = self.xl.Add(
                lo_srctype_s2i[source_type],
                source.api,
                None,
                yng_s2i[has_headers])
        else:
            xl = self.xl.Add(
                SourceType=lo_srctype_s2i[source_type],
                Source=source,
                LinkSource=True,
                XlListObjectHasHeaders=yng_s2i[has_headers],
                Destination=destination.api)
        return self._wrap(xl)


def _attr_listobjects(obj):
    return _attr_object_impl(ListObjects, obj.xl.ListObjects)


# --- Base of ListRowColumn ---
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

    @property
    def range(self):
        return xlwindows.Range(xl=self.xl.Range)


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


# --- ListColumn ---
class ListColumn(BaseListRowColumn):
    @property
    def name(self):
        return self.xl.Name

    @name.setter
    def name(self, value):
        self.xl.Name = value

    @property
    def totals_calculation(self):
        return totals_calc_i2s[self.xl.TotalsCalculation]

    @totals_calculation.setter
    def totals_calculation(self, calculation):
        self.xl.TotalsCalculation = totals_calc_s2i[calculation]

    @property
    def body(self):
        return xlwindows.Range(xl=self.xl.DataBodyRange)

    @property
    def total(self):
        return xlwindows.Range(xl=self.xl.Total)


class ListColumns(BaseListRowsColumns):
    _wrap = ListColumn


# --- QueryTable ---
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

    @property
    def adjust_column_width(self):
        return self.xl.AdjustColumnWidth

    @adjust_column_width.setter
    def adjust_column_width(self, adjust):
        self.xl.AdjustColumnWidth = adjust


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
    return _attr_object_impl(QueryTables, obj.xl.QueryTables)


# --- PageSetup ---
class PageSetup(object):
    def __init__(self, xl):
        self.xl = xl
        self._xlapp = self.xl.Application

    def __enter__(self):
        """
        Set the PrintCommunication property to False to speed up
        the execution of code that sets PageSetup properties.
        """
        self._xlapp.PrintCommunication = False
        return self

    def __exit__(self, exception_type, exception_value, traceback):
        """
        Set the PrintCommunication property to True after
        setting properties to commit all cached PageSetup commands.
        """
        self._xlapp.PrintCommunication = True

    def inches2pts(self, inches):
        return self._xlapp.InchesToPoints(inches)

    def cms2pts(self, cms):
        return self.CentimetersToPoints(cms)

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        """
        Returns the parent of the object.
        """
        return xlwindows.Sheet(self.xl.Parent)

    @property
    def fit_to_tall(self):
        return self.xl.FitToPagesTall

    @fit_to_tall.setter
    def fit_to_tall(self, pages):
        self.xl.FitToPagesTall = pages

    @property
    def fit_to_wide(self):
        return self.xl.FitToPagesWide

    @fit_to_wide.setter
    def fit_to_wide(self, pages):
        self.xl.FitToPagesWide = pages

    @property
    def zoom(self):
        return self.xl.Zoom

    @zoom.setter
    def zoom(self, ratio):
        self.xl.Zoom = ratio

    @property
    def orientation(self):
        return pg_ornt_i2s[self.xl.Orientation]

    @orientation.setter
    def orientation(self, aspect):
        self.xl.Orientation = pg_ornt_s2i[aspect]

    @property
    def paper_size(self):
        return paper_i2s[self.xl.PaperSize]

    @paper_size.setter
    def paper_size(self, size):
        self.xl.PaperSize = paper_s2i[size]

    @property
    def header_right(self):
        return self.xl.RightHeader

    @header_right.setter
    def header_right(self, text):
        self.xl.RightHeader = text

    @property
    def header_left(self):
        return self.xl.LeftHeader

    @header_left.setter
    def header_left(self, text):
        self.xl.LeftHeader = text

    @property
    def header_center(self):
        return self.xl.CenterHeader

    @header_center.setter
    def header_center(self, text):
        self.xl.CenterHeader = text

    @property
    def footer_right(self):
        return self.xl.RightFooter

    @footer_right.setter
    def footer_right(self, text):
        self.xl.RightFooter = text

    @property
    def footer_left(self):
        return self.xl.LeftFooter

    @footer_left.setter
    def footer_left(self, text):
        self.xl.LeftFooter = text

    @property
    def footer_center(self):
        return self.xl.CenterFooter

    @footer_center.setter
    def footer_center(self, text):
        self.xl.CenterFooter = text

    @property
    def margin_top(self):
        return self.xl.TopMargin

    @margin_top.setter
    def margin_top(self, pts):
        self.xl.TopMargin = pts

    @property
    def margin_bottom(self):
        return self.xl.BottomMargin

    @margin_bottom.setter
    def margin_bottom(self, pts):
        self.xl.BottomMargin = pts

    @property
    def margin_right(self):
        return self.xl.RightMargin

    @margin_right.setter
    def margin_right(self, pts):
        self.xl.RightMargin = pts

    @property
    def margin_left(self):
        return self.xl.LeftMargin

    @margin_left.setter
    def margin_left(self, pts):
        self.xl.LeftMargin = pts


def _attr_pagesetup(obj):
    return _attr_object_impl(PageSetup, obj.xl.PageSetup)


# --- Border ---
class Border(object):
    def __init__(self, xl):
        self.xl = xl

    @property
    def api(self):
        return self.xl

    @property
    def parent(self):
        return xlwindows.Range(xl=self.xl.Parent)

    @property
    def weight(self):
        return bd_wt_i2s[self.xl.Weight]

    @weight.setter
    def weight(self, wt):
        self.xl.Weight = bd_wt_s2i[wt]

    @property
    def style(self):
        return line_style_i2s[self.xl.LineStyle]

    @style.setter
    def style(self, style):
        self.xl.LineStyle = line_style_s2i[style]


class Borders(xlwindows.Collection):
    _wrap = Border

    @property
    def parent(self):
        return xlwindows.Range(xl=self.xl.Parent)

    def __call__(self, bds_index):
        return Border(xl=self.xl(bds_index_s2i[bds_index]))


def _attr_borders(obj):
    return _attr_object_impl(Borders, obj.xl.Borders)


# --- constants ---
totals_calc_s2i = {
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

totals_calc_i2s = {v: k for k, v in totals_calc_s2i.items()}

lo_srctype_s2i = {
    'external': 0,  # ListObjectSourceType.xlSrcExternal
    'query': 3,  # ListObjectSourceType.xlSrcQuery
    'range': 1,  # ListObjectSourceType.xlSrcRange
    'xml': 2  # ListObjectSourceType.xlSrcXml
    }

lo_srctype_i2s = {v: k for k, v in lo_srctype_s2i.items()}

yng_s2i = {
    'guess': 0,  # YesNoGuess.xlGuess
    'no': 2,  # YesNoGuess.xlNo
    'yes': 1  # YesNoGuess.xlYes
    }

yng_i2s = {v: k for k, v in yng_s2i.items()}

pg_ornt_s2i = {
    'landscape': 2,  # PageOrientation.xlLandscape
    'portrait': 1  # PageOrientation.xlPortrait
    }

pg_ornt_i2s = {v: k for k, v in pg_ornt_s2i.items()}

paper_s2i = {
    '10x14': 16,  # PaperSize.xlPaper10x14
    '11x17': 17,  # PaperSize.xlPaper11x17
    'a3': 8,  # PaperSize.xlPaperA3
    'a4': 9,  # PaperSize.xlPaperA4
    'a4_small': 10,  # PaperSize.xlPaperA4Small
    'a5': 11,  # PaperSize.xlPaperA5
    'b4': 12,  # PaperSize.xlPaperB4
    'b5': 13,  # PaperSize.xlPaperB5
    'c_sheet': 24,  # PaperSize.xlPaperCsheet
    'd_sheet': 25,  # PaperSize.xlPaperDsheet
    'envelope_10': 20,  # PaperSize.xlPaperEnvelope10
    'envelope_11': 21,  # PaperSize.xlPaperEnvelope11
    'envelope_12': 22,  # PaperSize.xlPaperEnvelope12
    'envelope_14': 23,  # PaperSize.xlPaperEnvelope14
    'envelope_9': 19,  # PaperSize.xlPaperEnvelope9
    'envelope_b4': 33,  # PaperSize.xlPaperEnvelopeB4
    'envelope_b5': 34,  # PaperSize.xlPaperEnvelopeB5
    'envelope_b6': 35,  # PaperSize.xlPaperEnvelopeB6
    'envelope_c3': 29,  # PaperSize.xlPaperEnvelopeC3
    'envelope_c4': 30,  # PaperSize.xlPaperEnvelopeC4
    'envelope_c5': 28,  # PaperSize.xlPaperEnvelopeC5
    'envelope_c6': 31,  # PaperSize.xlPaperEnvelopeC6
    'envelope_c65': 32,  # PaperSize.xlPaperEnvelopeC65
    'envelope_dl': 27,  # PaperSize.xlPaperEnvelopeDL
    'envelope_italy': 36,  # PaperSize.xlPaperEnvelopeItaly
    'envelope_monarch': 37,  # PaperSize.xlPaperEnvelopeMonarch
    'envelope_personal': 38,  # PaperSize.xlPaperEnvelopePersonal
    'e_sheet': 26,  # PaperSize.xlPaperEsheet
    'executive': 7,  # PaperSize.xlPaperExecutive
    'fanfold_legal_german': 41,  # PaperSize.xlPaperFanfoldLegalGerman
    'fanfold_std_german': 40,  # PaperSize.xlPaperFanfoldStdGerman
    'fanfold_us': 39,  # PaperSize.xlPaperFanfoldUS
    'folio': 14,  # PaperSize.xlPaperFolio
    'ledger': 4,  # PaperSize.xlPaperLedger
    'legal': 5,  # PaperSize.xlPaperLegal
    'letter': 1,  # PaperSize.xlPaperLetter
    'letter_small': 2,  # PaperSize.xlPaperLetterSmall
    'note': 18,  # PaperSize.xlPaperNote
    'quarto': 15,  # PaperSize.xlPaperQuarto
    'statement': 6,  # PaperSize.xlPaperStatement
    'tabloid': 3,  # PaperSize.xlPaperTabloid
    'user': 256,  # PaperSize.xlPaperUser
    }

paper_i2s = {v: k for k, v in paper_s2i.items()}

bds_index_s2i = {
    'left': 7,  # BordersIndex.xlEdgeLeft
    'right': 10,  # BordersIndex.xlEdgeRight
    'top': 8,  # BordersIndex.xlEdgeTop
    'bottom': 9,  # BordersIndex.xlEdgeBottom
    'diagonal_down': 5,  # BordersIndex.xlDiagonalDown
    'diagonal_up': 6  # BordersIndex.xlDiagonalUp
    }

bds_index_i2s = {v: k for k, v in bds_index_s2i.items()}

bd_wt_s2i = {
    'hairline': 1,  # BorderWeight.xlHairline
    'medium': -4138,  # BorderWeight.xlMedium
    'thick': 4,  # BorderWeight.xlThick
    'thin': 2  # BorderWeight.xlThin
    }

bd_wt_i2s = {v: k for k, v in bd_wt_s2i.items()}

line_style_s2i = {
    'continuous': 1,  # LineStyle.xlContinuous
    'dash': -4115,  # LineStyle.xlDash
    'dash_dot': 4,  # LineStyle.xlDashDot
    'dash_dotdot': 5,  # LineStyle.xlDashDotDot
    'dot': -4118,  # LineStyle.xlDot
    'double': -4119,  # LineStyle.xlDouble
    'none': -4142,  # LineStyle.xlLineStyleNone
    'slant_dashdot': 13,  # LineStyle.xlSlantDashDot
    }

line_style_i2s = {v: k for k, v in line_style_s2i.items()}
