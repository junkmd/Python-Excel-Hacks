import xlwings.main as xlmain
import sys

if sys.platform.startswith('win'):
    from . import _xlwindows_hacks as xlplatform_hacks
else:
    pass
    # from . import _xlmac_hacks as xlplatform_hacks
    # not yet implemented.


# --- Base of ListObject and QueryTable ---
class BaseTable(object):
    """
    internal class for ListObject and QueryTable.
    """
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
        return Sheet_Hacked(impl=self.impl.parent)

    @property
    def name(self):
        """
        Returns or sets the name of the object.
        """
        return self.impl.name

    def refresh(self):
        """
        Updates an external data range.
        For only based on the results of a SQL query.
        """
        self.impl.refresh()


class BaseTables(xlmain.Collection):
    """
    internal class for ListObjects and QueryTables.
    """
    @property
    def parent(self):
        """
        Returns the parent of the object.
        """
        return Sheet_Hacked(impl=self.impl.parent)

    @property
    def api(self):
        """
        Returns the native object (``pywin32`` or ``appscript`` obj)
        of the engine being used.
        """
        return self.impl.api


# --- ListObject ---
class ListObject(BaseTable):
    """
    Represents a ListObject object.
    An Object in the ListObjects collection.
    """
    @property
    def querytable(self):
        """
        Returns the QueryTable object that provides a link
        for the ListObject object to the list server.
        """
        return QueryTable(self.impl.querytable)

    @property
    def showtotals(self):
        """
        Gets or sets whether the Total row is visible.
        """
        return self.impl.showtotals

    @showtotals.setter
    def showtotals(self, value):
        self.impl.showtotals = value

    @property
    def listcolumns(self):
        """
        Returns a ListColumns collection that represents
        all the columns in a ListObject object.
        """
        return ListColumns(self.impl.listcolumns)

    def unlink(self):
        """
        Removes the link to a DB et al.
        """
        self.impl.unlink()

    @property
    def range(self):
        """
        Returns a Range object that represents the range to which
        the specified list object in the above list applies.
        """
        return xlmain.Range(impl=self.impl.range)

    @property
    def header_row(self):
        """
        Returns a Range object that represents the range of
        the header row for a list.
        """
        return xlmain.Range(impl=self.impl.header_row)

    @property
    def body(self):
        """
        Returns a Range object that represents the range of values,
        excluding the header row, in a table.
        """
        return xlmain.Range(impl=self.impl.body)

    @property
    def totals_row(self):
        """
        Returns a Range representing the Total row.
        """
        return xlmain.Range(impl=self.impl.totals_row)


class ListObjects(BaseTables):
    """
    A collection of all the ListObject objects on a worksheet.
    Each ListObject object represents a table in the worksheet.
    """
    _wrap = ListObject

    def add(self, source_type, source, destination=None, has_headers='guess'):
        """
        Creates a new list object.

        Parameters
        ----------
        source_type : 'external', 'query', 'range' or 'xml'
            Indicates the kind of source for the query.
        source : str or Range.
            If source_type was 'range', must be Range.
        destination : Range or None.
            If source_type was 'range', must be None.
        has_headers : 'yes', 'no' or default 'guess'
            'guess' : Excel determines whether there is a header.
            'yes' : Top row of range will be header.
            'no' : The header row will be added to top of the entire range.

        -------
        """
        if destination is not None:
            dest = destination.impl
        else:
            dest = None

        if source_type == 'range':
            src = source.impl
        else:
            src = source

        return ListObject(
            self.impl.add(
                source_type, src, dest, has_headers)
            )


# --- Base of ListRowColumn ---
class BaseListRowColumn(object):
    """internal class."""
    def __init__(self, impl):
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
        return ListObject(impl=self.impl.parent)

    @property
    def range(self):
        return xlmain.Range(impl=self.impl.range)


class BaseListRowsColumns(xlmain.Collection):
    """
    internal class.
    _wrap attribute must be not None.
    """
    def __init__(self, impl):
        xlmain.Collection.__init__(self, impl)

    @property
    def parent(self):
        """
        Returns the parent of the object.
        """
        return ListObject(impl=self.impl.parent)


# --- ListColumn ---
class ListColumn(BaseListRowColumn):
    """
    Represents a column in a table.
    """
    @property
    def name(self):
        return self.impl.name

    @name.setter
    def name(self, value):
        self.impl.name = value

    @property
    def totals_calculation(self):
        """
        Gets or sets determineing the type of calculation in the Totals row.
        """
        return self.impl.totals_calculation

    @totals_calculation.setter
    def totals_calculation(self, calculation):
        self.impl.totals_calculation = calculation

    @property
    def body(self):
        """
        Returns a Range object that is the size of
        the data portion of a column.
        """
        return xlmain.Range(impl=self.impl.body)

    @property
    def total(self):
        """
        Returns the Total row.
        """
        return xlmain.Range(impl=self.impl.total)


class ListColumns(BaseListRowsColumns):
    """
    A collection of all the ListColumn objects in the specified ListObject.
    """
    _wrap = ListColumn


# --- QueryTable ---
class QueryTable(BaseTable):
    """
    Represents a QueryTable object.
    An Object in the QueryTables collection.
    """
    @property
    def background_query(self):
        """
        Gets or sets the performance of refreshing to True or False.

            True: The query table are performed asynchronously.
            False: The query table are NOT performed asynchronously.
        """
        return self.impl.background_query

    @background_query.setter
    def background_query(self, value):
        self.impl.background_query = value

    @property
    def command_text(self):
        """
        Returns or sets the command string for the data source.
        """
        return self.impl.command_text

    @command_text.setter
    def command_text(self, text):
        self.impl.command_text = text

    @property
    def listobject(self):
        """
        Returns a ListObject object for the QueryTable object.
        """
        return ListObject(self.impl.listobject)

    @property
    def adjust_column_width(self):
        return self.impl.adjust_column_width

    @adjust_column_width.setter
    def adjust_column_width(self, adjust):
        self.impl.adjust_column_width = adjust


class QueryTables(BaseTables):
    """
    A collection of all QueryTable objects on a worksheet.
    Each QueryTable object represents a table in the worksheet.

    Examples
    --------

    .. code-block:: python

        import xlwings as xw

        # now rewriting...
    """
    _wrap = QueryTable

    def add(self, connection, destination, sql=None):
        """
        Creates a new QueryTable.

        Parameters
        ----------
        connection : str, ADO/DAO recordset, web query, data finder, text file
            A datasource of the table.
        destination : Range
            A range in the upper-left corner of the Sheet.
        sql : str, default None
            A SQL query str.

        -------
        """
        impl = self.impl.add(
            connection,
            destination.impl,
            sql)
        return self._wrap(impl)


# --- PageSetup ---
class PageSetup(object):
    """
    Represents the page setup description.
    """
    def __init__(self, impl=None):
        self.impl = impl

    def __enter__(self):
        self.impl.__enter__()
        return self

    def __exit__(self, exception_type, exception_value, traceback):
        self.impl.__exit__(exception_type, exception_value, traceback)

    def inches2pts(self, inches):
        """
        Converts a measurement from inches to points.
        """
        return self.impl.inches2pts(inches)

    def cms2pts(self, cms):
        """
        Converts a measurement from centimeters to points.
        """
        return self.impl.cms2pts(cms)

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
        return Sheet_Hacked(impl=self.impl.parent)

    @property
    def fit_to_tall(self):
        """
        Returns or sets the number of pages tall the worksheet
        will be scaled to when it's printed.
        If the zoom property is True , this is ignored.
        """
        return self.impl.fit_to_tall

    @fit_to_tall.setter
    def fit_to_tall(self, pages):
        self.impl.fit_to_tall = pages

    @property
    def fit_to_wide(self):
        """
        Returns or sets the number of pages wide the worksheet
        will be scaled to when it's printed.
        If the zoom property is True , this is ignored.
        """
        return self.impl.fit_to_wide

    @fit_to_wide.setter
    def fit_to_wide(self, pages):
        self.impl.fit_to_wide = pages

    @property
    def zoom(self):
        """
        Returns or sets a percentage(between 10 and 400 percent) or False that
        represents scale of the worksheet for printing.
        """
        return self.impl.zoom

    @zoom.setter
    def zoom(self, ratio):
        self.impl.zoom = ratio

    @property
    def orientation(self):
        """
        Returns or sets a 'portrait'(taller) or 'landscape'(wider) that
        represents the printing mode.
        """
        return self.impl.orientation

    @orientation.setter
    def orientation(self, aspect):
        self.impl.orientation = aspect

    @property
    def paper_size(self):
        """
        Returns or sets the size of the paper.
        """
        return self.impl.paper_size

    @paper_size.setter
    def paper_size(self, size):
        self.impl.paper_size = size

    @property
    def header_right(self):
        """
        Returns or sets the text of header.
        """
        return self.impl.header_right

    @header_right.setter
    def header_right(self, text):
        self.impl.header_right = text

    @property
    def header_left(self):
        """
        Returns or sets the text of header.
        """
        return self.impl.header_left

    @header_left.setter
    def header_left(self, text):
        self.impl.header_left = text

    @property
    def header_center(self):
        """
        Returns or sets the text of header.
        """
        return self.impl.header_center

    @header_center.setter
    def header_center(self, text):
        self.impl.header_center = text

    @property
    def footer_right(self):
        """
        Returns or sets the text of footer.
        """
        return self.impl.footer_right

    @footer_right.setter
    def footer_right(self, text):
        self.impl.footer_right = text

    @property
    def footer_left(self):
        """
        Returns or sets the text of footer.
        """
        return self.impl.footer_left

    @footer_left.setter
    def footer_left(self, text):
        self.impl.footer_left = text

    @property
    def footer_center(self):
        """
        Returns or sets the text of footer.
        """
        return self.impl.footer_center

    @footer_center.setter
    def footer_center(self, text):
        self.impl.footer_center = text

    @property
    def margin_top(self):
        """
        Returns or sets the size of the margin.
        """
        return self.impl.margin_top

    @margin_top.setter
    def margin_top(self, pts):
        self.impl.margin_top = pts

    @property
    def margin_bottom(self):
        """
        Returns or sets the size of the margin.
        """
        return self.impl.margin_bottom

    @margin_bottom.setter
    def margin_bottom(self, pts):
        self.impl.margin_bottom = pts

    @property
    def margin_right(self):
        """
        Returns or sets the size of the margin.
        """
        return self.impl.margin_right

    @margin_right.setter
    def margin_right(self, pts):
        self.impl.margin_right = pts

    @property
    def margin_left(self):
        """
        Returns or sets the size of the margin.
        """
        return self.impl.margin_left

    @margin_left.setter
    def margin_left(self, pts):
        self.impl.margin_left = pts


# --- implemented sheet ---
class Sheet_Hacked(xlmain.Sheet):
    """
    Hacked xlwings.main.Sheet.

    Examples
    --------

    .. code-block:: python

        import xlwings as xw

        wb = xw.Book()
        ws = Sheet_Hacked(impl=wb.sheets.add().impl)
        # or
        # ws = Sheet_Hacked(impl=wb.sheets.add[0].impl)
    ----------

    """
    def __init__(self, impl):
        """
        Construct an Sheet object with the extra properties.

        Parameters
        ----------
        impl : xlwings.Sheet.impl or xlwings.main.xlplatform.Sheet

        """
        if not isinstance(impl, xlmain.xlplatform.Sheet):
            raise TypeError(
                "'impl' must be instance of xlwings.Sheet.impl")
        xlmain.Sheet.__init__(self, impl=impl)

    @property
    def listobjects(self):
        """
        A collection of all the ListObject objects on a worksheet.
        Each ListObject object represents a table in the worksheet.
        """
        return ListObjects(
            impl=xlplatform_hacks._attr_listobjects(self.impl)
        )

    @property
    def querytables(self):
        """
        Represents a worksheet table built from data returned from
        an external data source,
        such as an SQL server or a Microsoft Access database.
        """
        return QueryTables(
            impl=xlplatform_hacks._attr_querytables(self.impl)
        )

    @property
    def pagesetup(self):
        """
        Contains all page setup attributes (left margin, bottom margin,
        paper size, and so on) as properties.

        Suspends communication with the printer within the 'with' block.

        Examples
        --------

        .. code-block:: python

            import xlwings as xw

            ws = Sheet_Hacked(xw.Sheet())

            with ws.pagesetup as psu:
                psu.fit_to_tall = 1

        """
        return PageSetup(
            impl=xlplatform_hacks._attr_pagesetup(self.impl)
        )


# --- Border ---
class Border(object):
    def __init__(self, impl):
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
        return xlmain.Range(impl=self.impl.parent)

    @property
    def weight(self):
        """
        Returns or sets value that represents the weight of the border.
        'hairline', 'medium', 'thick' or 'thin'.
        """
        return self.impl.weight

    @weight.setter
    def weight(self, wt):
        self.impl.weight = wt

    @property
    def style(self):
        """
        Returns or sets the line style for the border.
        'continuous', 'dash', 'dash_dot', 'dash_dotdot', 'dot', 'double',
        'none', 'slant_dashdot' or None.
        """
        return self.impl.style

    @style.setter
    def style(self, style):
        self.impl.style = style

    @property
    def color(self):
        """
        Returns or sets the primary color of the object.
        """
        return self.impl.color

    @color.setter
    def color(self, color_or_rgb):
        self.impl.color = color_or_rgb

    @property
    def tint_and_shade(self):
        """
        Returns or sets lightens or darkens of border color.
        From -1(darkest) to 1(lightest), and Zero(0) is neutral.
        """
        return self.impl.tint_and_shade

    @tint_and_shade.setter
    def tint_and_shade(self, single):
        self.impl.tint_and_shade = single


class Borders(xlmain.Collection):
    """
    A collection of six Border objects.
    Order of borders are;
    'left', 'right', 'top', 'bottom', 'diagonal_down', 'diagonal_up'
    """
    _wrap = Border

    @property
    def parent(self):
        """
        Returns the parent of the object.
        """
        return xlmain.Range(xl=self.impl.parent)


def get_borders_of(range):
    """
    Returns a collection of six Border objects.
    Order of borders are;
    'left', 'right', 'top', 'bottom', 'diagonal_down', 'diagonal_up'
    """
    return Borders(xlplatform_hacks._attr_borders(range.impl))
