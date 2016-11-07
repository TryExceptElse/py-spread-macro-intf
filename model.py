"""
Handles macro interaction with a spreadsheet workbook
"""

try:
    import xlwings as xw
except ImportError:
    xw = None


DEFAULT_COLOR = -1
MAX_CELL_GAP = 10


class Model:
    """
    Abstract model to be extended by office interface specific
    subclasses in Office.
    """
    def __init__(self) -> None:
        raise NotImplementedError

    def __getitem__(self, item: str or int):
        """
        Gets sheet from model, either by the str name of the sheet,
        or the int index.
        :param item: str or int
        :return: Office.Sheet
        """
        raise NotImplementedError
        # implemented by office program specific subclasses

    def __iter__(self):
        raise NotImplementedError

    def get_sheet(
            self,
            sheet_name: str,
            row_ref_i: int=0,
            col_ref_i: int=0
    ):
        """
        Gets sheet of passed name in Model.
        Functions the same as Model.__getitem__ except reference row
        and reference column indices may be passed to this method.
        :param sheet_name: str
        :param row_ref_i: int
        :param col_ref_i: int
        :return: Sheet
        """
        raise NotImplementedError  # todo: finish sorting these two methods out

    def sheet_exists(self, *sheet_name: str) -> str:
        raise NotImplementedError
        # implemented by office program specific subclasses

    @property
    def sheets(self):
        raise NotImplementedError
        # implemented by office program specific subclasses

    @property
    def sheet_names(self):
        """
        Gets iterable of sheet names
        :return: str iterable
        """
        raise NotImplementedError


class Sheet:
    i7e_sheet = None  # interface sheet obj. ie; com.sun.star...Sheet
    _reference_row_index = 0
    _reference_column_index = 0

    def __init__(
            self,
            i7e_sheet,
            reference_row_index=0,
            reference_column_index=0
    ) -> None:
        raise NotImplementedError

    def get_column(
            self,
            column_identifier: int or float or str
    ):
        """
        Gets column by name if identifier is str, otherwise,
        attempts to get column by index.
        :param column_identifier: int, float, or str
        :return: Office.Column
        """
        if isinstance(column_identifier, str):
            return self.get_column_by_name(column_identifier)
        else:
            return self.get_column_by_index(column_identifier)

    def get_column_by_index(self, column_index: int):
        raise NotImplementedError
        # implemented by office program specific subclasses

    def get_column_by_name(self, column_name: int or float or str):
        """
        Gets column from a passed reference value which is compared
        to each cell value in the reference row.
        This function will return the first column whose name matches
        the passed value.
        :return: Office.Column
        """
        x = self.get_column_index_from_name(column_name)
        return self.get_column_by_index(x) if x is not None else None

    def get_column_index_from_name(
        self,
        column_name: int or float or str
    ) -> int or None:
        """
        Gets column index from name
        :param column_name: int, float, or str
        :return: int or None
        """
        for x, cell in enumerate(self.reference_row):
            if cell.value == column_name:
                return x

    def get_row(self, row_identifier: int or float or str):
        """
        Gets row by name if identifier is str, otherwise by index
        :param row_identifier: int, float, or str
        :return: Office.Row
        """
        if isinstance(row_identifier, str):
            return self.get_row_by_name(row_identifier)
        else:
            return self.get_row_by_index(row_identifier)

    def get_row_by_index(self, row_index: int or str):
        raise NotImplementedError
        # implemented by office program specific subclasses

    def get_row_by_name(self, row_name: int or str or float):
        """
        Gets row from a passed reference value which is compared
        to each cell value in the reference row.
        This function will return the first row whose name matches
        the passed value.
        :return: Office.Column
        """
        y = self.get_row_index_from_name(row_name)
        return self.get_row_by_index(y) if y is not None else None

    def get_row_index_from_name(
            self,
            row_name: int or float or str
    ) -> int or None:
        """
        Gets index of a row from passed name
        :param row_name: int, float, or str
        :return: int or None
        """
        for y, cell in enumerate(self.rows):
            if cell.value == row_name:
                return y

    def get_cell(self, cell_identifier, **kwargs):
        # if cell_identifier is list or tuple, get cell via
        # x, y coordinates
        if isinstance(cell_identifier, (list, tuple)):
            x_identifier = cell_identifier[0]
            y_identifier = cell_identifier[1]
            x_identifier_type = kwargs.get('x_identifier_type', None)
            y_identifier_type = kwargs.get('y_identifier_type', None)
            # sanity check
            assert x_identifier_type in ('index', 'name', None), \
                "x identifier type should be 'index', 'name', or None." \
                "Instead got %s" % x_identifier_type
            assert y_identifier_type in ('index', 'name', None), \
                "y identifier type should be 'index', 'name', or None." \
                'instead got %s' % y_identifier_type
            if (x_identifier_type == 'name' or
                x_identifier_type is None and isinstance(
                    x_identifier, str)):
                column = self.get_column_by_name(x_identifier)
            else:
                assert isinstance(x_identifier, int)  # sanity check
                column = self.get_column_by_index(x_identifier)
            if (y_identifier_type == 'name' or
                x_identifier_type is None and isinstance(
                    y_identifier, str)):
                return column.get_cell_by_reference(y_identifier)
            else:
                return column.get_cell_by_index(y_identifier)

    @property
    def reference_row_index(self) -> int:
        """
        Gets index of reference row
        :return: int
        """
        return self._reference_row_index

    @reference_row_index.setter
    def reference_row_index(self, new_index: int) -> None:
        """
        Sets reference row by passing the index of the new row.
        Must be > 0
        :param new_index: int
        :return: None
        """
        if new_index < 0:
            raise IndexError('Reference row index must be > 0')
        self._reference_row_index = new_index

    @property
    def reference_row(self):
        """
        Gets reference row
        :return: Office.Row
        """
        return self.get_row(self.reference_row_index)

    @property
    def reference_column_index(self) -> int:
        """
        Gets index of reference column
        :return: int
        """
        return self._reference_column_index

    @reference_column_index.setter
    def reference_column_index(self, new_index) -> None:
        """
        Sets reference column by passing the index of the new
        reference column.
        Must be > 0.
        :param new_index: int
        :return: None
        """
        if new_index < 0:
            raise IndexError('Reference row index must be > 0')
        self._reference_column_index = new_index

    @property
    def reference_column(self):
        """
        Gets reference column
        :return: Office.Column
        """
        return self.get_column(self.reference_column_index)

    @property
    def columns(self):
        return LineSeries(reference_line=self.reference_row)

    @property
    def rows(self):
        return LineSeries(reference_line=self.reference_column)

    def __str__(self) -> str:
        raise NotImplementedError


class LineSeries:
    """Class storing collection of Line, Column, or Row objects"""

    def __init__(self, reference_line) -> None:
        if not isinstance(reference_line, Line):
            raise TypeError('Expected LineSeries to be passed reference line.'
                            'Got instead: %s' % repr(reference_line))
        self.reference_line = reference_line

    def __getitem__(self, item: int or float or str):
        """
        If item is int, returns line of that index, otherwise looks
        for a line of that name.
        :param item: int, float, or str
        :return: Line or None
        """
        if isinstance(item, int):
            return self.get_by_index(item)
        else:
            return self.get_by_name(item)

    def __iter__(self):
        """
        Returns Generator that iterates over columns in LineSeries
        :return: Generator<Line>
        """
        for cell in self.reference_line:
            if self._contents_type == 'rows':
                yield cell.row
            elif self._contents_type == 'columns':
                yield cell.column

    def __len__(self):
        """
        Returns size of LineSeries
        :return: int
        """
        count = 0
        try:
            while self.__iter__().__next__():
                count += 1
        except StopIteration:
            return count

    def get_by_name(self, name: int or float or str):
        """
        Gets line from passed line name.
        Returns None if no line of that name is found.
        :param name: int, float or str
        :return: Line or None
        """
        for cell in self.reference_line:
            if cell.value == name:
                if self._contents_type == 'columns':
                    return cell.column
                elif self._contents_type == 'rows':
                    return cell.row

    def get_by_index(self, index: int):
        """
        Gets line from passed line index
        :param index:
        :return: Line
        """
        if self._contents_type == 'rows':
            return self.sheet.get_row_by_index(index)
        elif self._contents_type == 'columns':
            return self.sheet.get_column_by_index(index)

    @property
    def sheet(self) -> Sheet:
        """
        Gets sheet that owns the Columns or Rows of this Lines obj.
        :return: Sheet
        """
        return self.reference_line.sheet

    @property
    def names(self):
        """
        Yields names of lines in LineList
        :return: int, float, or str
        """
        for line in self:
            yield line.name

    @property
    def named_only(self):
        """
        Yields only lines in this series that have line headers
        :return:
        """
        for line in self:
            if line.name:
                yield line

    @property
    def indexes(self):
        """
        Yields indexes of lines in LineList
        :return: int
        """
        for line in self:
            yield line.index

    @property
    def _contents_type(self) -> str:
        if isinstance(self.reference_line, Row):
            return 'columns'
        elif isinstance(self.reference_line, Column):
            return 'rows'


class Line:
    sheet = None  # these are to be set on init in subclasses
    index = None  # index of this line.

    def __init__(
        self,
        sheet: Sheet,
        index: int,
        reference_index: int,
    ) -> None:
        if not isinstance(sheet, Sheet):
            raise TypeError(
                'Expected subclass of Sheet. Instead got %s'
                % repr(sheet))
        if not isinstance(index, int):
            raise TypeError(
                'Expected line index to be an int. '
                'Instead got %s' % repr(index))
        if not isinstance(reference_index, int):
            raise TypeError(
                'Expected reference name index to be an int, '
                'Instead got %s' % repr(reference_index))
        self.sheet = sheet
        self.index = index
        self.reference_index = index

    def __getitem__(self, item: int or str):
        raise NotImplementedError
        # implemented by office program specific subclasses

    def __iter__(self):
        raise NotImplementedError
        # implemented by office program specific subclasses

    def __len__(self):
        raise NotImplementedError
        # implemented by office program specific subclasses

    def get_cell_by_index(self, index: int):
        raise NotImplementedError
        # implemented by Row and Column in office program
        # specific subclasses

    def get_cell_by_reference(self, reference: str or float or int):
        for i, cell in enumerate(self._reference_line):
            if cell.value == reference:
                return self.get_cell_by_index(i)

    def get_iterator(self, axis: str):
        assert axis == 'x' or axis == 'y'
        return CellLine(self.sheet, axis, self.index)

    def clear(self, include_header: bool = False):
        """
        Clears line of cells.
        If Include header is True; clears cell data in cells
        preceding and including re
        a
        :param include_header: bool
        :return: None
        """
        [cell.clear() for i, cell in enumerate(self)
         if i > self.name_cell_index or include_header]

    @property
    def _reference_line(self):
        raise NotImplementedError

    @property
    def duplicates(self):
        """
        Returns generator of duplicate cells in Column.
        :return: cells (iterator)
        """
        cell_values = set()
        for cell in self:
            if cell.value_without_whitespace in cell_values:
                yield cell
            cell_values.add(cell.value_without_whitespace)

    @property
    def name_cell_index(self) -> int:
        """
        Gets the index of the cell which contains this line's name.
        :return: int
        """
        raise NotImplementedError

    @property
    def name(self):
        """
        Returns name of line, which is the value stored in the
        line's header cell, located in the sheet's reference
        row or column.
        :return: int, float, str or None
        """
        return self[self.name_cell_index].value

    def __repr__(self) -> str:
        return '%s(sheet=%s, index(0-base)=%s, ref_index=%s) name: %s' % (
            self.__class__.__name__,
            repr(self.sheet),
            self.index,
            self.reference_index,
            self.name
        )

    def __str__(self) -> str:
        return '%s[index(0-base): %s] name: %s' % (
            self.__class__.__name__,
            self.index,
            self.name
        )


class Column(Line):
    """
    Abstract Column class, extended by Office.XW.Column and Office.XW.Row
    """

    def __getitem__(self, cell_identifier):
        """
        Gets cell from passed identifier.
        If identifier is string, presumes it is a cell's name.
        If identifier is number, presumes it is a
        cell's index.
        To ensure the right method of fetching a cell is used,
        use .get_by_name or .get_by_index
        :param cell_identifier: str or int
        :return: Cell
        """
        assert isinstance(cell_identifier, (int, str)), \
            'Expected cell_identifier to be int or str, got %s ' \
            % cell_identifier
        if isinstance(cell_identifier, int):
            return self.get_cell_by_index(cell_identifier)
        else:
            for x, cell in enumerate(self.reference_column):
                if cell.value == cell_identifier:
                    return self[x]

    @property
    def _reference_line(self):
        return self.reference_column

    @property
    def reference_column(self):
        return self.sheet.reference_column

    @property
    def name_cell_index(self) -> int:
        """
        Gets the index of the cell which contains this
        Column's name.
        :return: int
        """
        return self.sheet.reference_row_index


class Row(Line):
    """
    Abstract Row obj. Extended by Office.XW.Row and Office.Uno.Row
    """

    def __getitem__(self, cell_identifier):
        """
        Gets cell from passed identifier.
        If identifier is string, presumes it is a cell's name.
        If identifier is number, presumes it is a
        cell's index.
        To ensure the right method of fetching a cell is used,
        use .get_by_name or .get_by_index.
        :param cell_identifier: str or int
        :return: Cell
        """
        assert isinstance(cell_identifier, (str, int))
        if isinstance(cell_identifier, int):
            return self.get_cell_by_index(cell_identifier)
        else:
            for x, cell in enumerate(self.reference_row):
                if cell.value == cell_identifier:
                    return self[x]

    @property
    def _reference_line(self):
        return self.reference_row

    @property
    def reference_row(self):
        """
        Gets reference row.
        This is the row that contains the names of the
        intersecting columns, allowing fetching of cells in
        this row via passage of a reference string
        :return: Row
        """
        return self.sheet.reference_row

    @property
    def name_cell_index(self) -> int:
        """
        Gets the index of the cell which contains this Row's name.
        :return: int
        """
        return self.sheet.reference_column_index


class Cell:
    position = None
    sheet = None

    def __init__(
            self,
            sheet: Sheet,
            position: tuple
    ) -> None:
        assert len(position) == 2
        assert all([isinstance(item, int) for item in position])
        self.position = tuple(position)
        self.sheet = sheet

    def set_color(self, color: int or list or tuple) -> None:
        raise NotImplementedError

    def get_color(self) -> int:
        raise NotImplementedError

    def remove_whitespace(self):
        self.value = self.value_without_whitespace

    def clear(self):
        """Clears cell by setting value to None and color to default"""
        self.value = None
        self.set_color(DEFAULT_COLOR)

    @property
    def row(self) -> Row:
        return self.sheet.get_row(self.y)

    @property
    def column(self) -> Column:
        return self.sheet.get_column(self.x)

    @property
    def has_whitespace(self) -> bool:
        """
        Gets bool of whether cell string contains whitespace.
        If cell contains a number, returns False.
        :return: bool
        """
        return self.value_without_whitespace != self.value

    @property
    def value_without_whitespace(self) -> str:
        """
        Gets value of cell without whitespace.
        If cell is not a string, returns value unchanged.
        :return:
        """
        if isinstance(self.value, str):
            return ' '.join(self.value.split())
        else:
            return self.value

    @property
    def value(self) -> int or float or str or None:
        raise NotImplementedError

    @value.setter
    def value(self, new_value: str or int or float or None) -> None:
        raise NotImplementedError

    @property
    def string(self):
        raise NotImplementedError

    @string.setter
    def string(self, new_string: str) -> None:
        raise NotImplementedError

    @property
    def float(self):
        raise NotImplementedError

    @float.setter
    def float(self, new_float: int or float) -> None:
        raise NotImplementedError

    @property
    def x(self):
        """
        Gets x position of cell
        :return: int
        """
        return self.position[0]

    @property
    def y(self) -> int:
        """
        Gets y position of cell
        :return: int
        """
        return self.position[1]

    def __repr__(self) -> str:
        return 'Cell(%s, %s) Value: %s' % (
            self.sheet, self.position, self.value.__repr__()
        )

    def __str__(self) -> str:
        return 'Cell[%s, Value: %s]' % (self.position, self.value.__repr__())


class CellLine:
    """
    Generator iterator that returns cells of a particular row or column
    """
    sheet = None
    axis = None
    index = None
    i = 0
    highest_inhabited_i = -1

    # max_i = 0

    def __init__(self, sheet: Sheet, axis: str, index: int) -> None:
        assert axis in ('x', 'y')
        if not isinstance(sheet, Sheet):
            raise TypeError('CellLine Constructor sheet arg should be a Sheet.'
                            ' Got instead: %s' % sheet.__repr__())
        if not isinstance(index, int):
            raise TypeError('CellLine __init__ index arg should be int. Got: '
                            '%s' % index.__repr__())
        self.sheet = sheet
        self.axis = axis
        self.index = index

    def __iter__(self):
        return self

    def __next__(self) -> Cell:
        # set starting x, y values
        x, y = (self.index, self.i) if self.axis == 'y' else \
            (self.i, self.index)
        cell = self.sheet.get_cell((x, y))  # get first cell
        # if cell is empty, look to see if a cell with a value follows.
        if cell.string == '' and self.i > self.highest_inhabited_i:
            for i in range(1, MAX_CELL_GAP):
                test_x, test_y = (self.index, self.i + i) if \
                    self.axis == 'y' else (self.i + i, self.index)
                test_cell = self.sheet.get_cell((test_x, test_y))
                if test_cell.string != '':
                    # if there is, mark that index as the highest i searched
                    # and break.
                    self.highest_inhabited_i = self.i + i
                    break
            else:
                # otherwise end iteration
                raise StopIteration()
        self.i += 1
        return cell


class Interface:
    """Abstract class inherited from by XW and Uno interfaces."""

    class Model:
        pass

    class Sheet:
        pass

    class Line:
        pass

    class Column:
        pass

    class Row:
        pass

    class Cell:
        pass


class Office:
    """
    Handles interface with workbook.

    This may not need to be a class, but any independent functions appear
    as their own macro, so this is instead its own class
    """

    class XW(Interface):
        """
        Handles XLWings interfacing with Office program
        """

        class Model(Model):
            def __init__(self, app_=None):
                if app_:
                    self.active_app = app_
                else:
                    try:
                        self.active_app = xw.apps[0]  # get first app open
                    except IndexError:
                        raise EnvironmentError(
                            'Office does not appear to have any running '
                            'instances.'
                        )
                # there should be only one open at a given time usually,
                # if any.

            def sheet_exists(self, *sheet_name: str) -> str:
                """
                Tests if sheet exists in any book.
                May be passed multiple sheet names.
                Returns first passed name that exists in books, or None
                :param sheet_name: str
                :return: str
                """
                for sheet_name_ in sheet_name:
                    if "::" in sheet_name_:
                        # if book/sheet name separator is in sheet_name,
                        # first find the book, then sheet
                        book_name, sheet_name_ = sheet_name_.split("::")
                        try:
                            book = self.active_app.books[book_name]
                            sheet = book.sheets[sheet_name_]
                        except KeyError:
                            continue
                        else:
                            assert sheet.name == sheet_name_
                            return sheet_name_
                    else:
                        # otherwise just find the sheet name
                        for sheet in self._xw_sheets:
                            if sheet_name_ == sheet.name:
                                return sheet_name_

            def __getitem__(self, item: str or int):
                """
                Gets passed item, returning the sheet of that name.
                :param item:
                :return: Sheet
                """
                if isinstance(item, str) and "::" in item:
                    # split and find book + name
                    book_name, sheet_name = item.split("::")
                    return Office.XW.Sheet(
                        self.active_app.books[book_name].sheets[sheet_name]
                    )
                else:
                    # otherwise just look everywhere
                    for sheet in self._xw_sheets:
                        if sheet.name == item:
                            return Office.XW.Sheet(
                                sheet
                            )

            @property
            def books(self):
                """
                Returns dict? of books open
                :return: dict? of books.
                """
                return self.active_app.books

            @property
            def _xw_sheets(self):
                """
                Yields each found xw sheet object in model
                :return: XW Sheet iterator
                """
                for xw_book in self.books:
                    for xw_sheet in xw_book.sheets:
                        yield xw_sheet

            @property
            def sheets(self):
                """
                Generator returning each sheet in Model.
                Weird implementation here to make it align with
                PyUno interface. This returns all sheets in all
                books open.
                :return: Sheet
                """
                for xw_sheet in self._xw_sheets:
                    yield Office.XW.Sheet(xw_sheet)

            @property
            def sheet_names(self):
                """
                Gets iterable of names of usable sheets in Model
                :return: iterator
                """
                for book in self.books:
                    for sheet in book.sheets:
                        yield "%s::%s" % (book.name, sheet.name)

        class Sheet(Sheet):
            def __init__(
                    self,
                    xw_sheet,
                    reference_row_index=0,
                    reference_column_index=0
            ) -> None:
                self.i7e_sheet = xw_sheet
                self.reference_row_index = reference_row_index
                self.reference_column_index = reference_column_index

            def get_row_by_index(self, row_index: int or str):
                if not isinstance(row_index, int):
                    raise TypeError('Row index must be an int, got %s'
                                    % row_index)
                if row_index < 0:
                    raise ValueError('Passed index must be 0 or greater.')
                return Office.XW.Row(self, row_index, self.reference_row_index)

            def get_column_by_index(self, column_index: int):
                if not isinstance(column_index, int):
                    raise TypeError('Column index must be an int, got %s'
                                    % column_index)
                if column_index < 0:
                    raise ValueError('Passed index must be 0 or greater.')
                return Office.XW.Column(
                    self,
                    column_index,
                    self.reference_row_index
                )

            def __str__(self) -> str:
                return 'Sheet[%s::%s]' % (
                    self.i7e_sheet.name,
                    self.i7e_sheet.book.name,
                )

        class Line(Line):
            def __len__(self) -> int:
                count = 0
                for each in self:
                    count += 1
                return count

        class Column(Line, Column):
            def __init__(
                    self,
                    sheet: Sheet,
                    column_index: int,
                    reference_column_index: int=0):
                super().__init__(
                    sheet=sheet,
                    index=column_index,
                    reference_index=reference_column_index
                )

            def get_cell_by_index(self, index: int):
                if not isinstance(index, int):
                    raise TypeError(
                        'get_cell_by_index: passed non-int index: %s' % index
                    )
                if index < 0:
                    raise ValueError(
                        'get_cell_by_index: passed index is < 0: %s' % index
                    )
                return Office.XW.Cell(self.sheet, (self.index, index))

            def __iter__(self):
                return self.get_iterator(axis='y')

        class Row(Line, Row):
            def __init__(
                    self,
                    sheet: Sheet,
                    row_index: int,
                    reference_row_index: int=0
            ) -> None:
                super().__init__(
                    sheet=sheet,
                    index=row_index,
                    reference_index=reference_row_index,
                )

            def get_cell_by_index(self, index: int):
                if not isinstance(index, int):
                    raise TypeError(
                        'get_cell_by_index: passed non-int index: %s' % index
                    )
                if index < 0:
                    raise ValueError(
                        'get_cell_by_index: passed index is < 0: %s' % index
                    )
                return Office.XW.Cell(self.sheet, (index, self.index))

            def __iter__(self):
                return self.get_iterator(axis='x')

        class Cell(Cell):
            def set_color(self, color: int or list or tuple) -> None:
                if color >= 0:
                    color = Color(color)
                    self._range.color = color.rgb
                elif color == -1:
                    self._range.color = None

            def get_color(self) -> int:
                color_int = self._range.color
                if color_int is None:
                    color_int = -1
                return color_int

            @property
            def _range(self):
                """
                Gets XW Range obj for this cell
                :return: xlwings.Range
                """
                x, y = self.position
                x += 1
                y += 1  # correct to excel 1 based index
                # XW passes position tuples as row, column
                return self.sheet.i7e_sheet.range(y, x)

            @property
            def value(self) -> int or float or str or None:
                return self._range.value

            @value.setter
            def value(self, new_value) -> None:
                self._range.value = new_value

            @property
            def float(self):
                # XW will only return number (float), str or None in most
                # cases, others include Date, (etc)?
                if isinstance(self.value, float):
                    return self.value
                else:
                    return 0.

            @property
            def string(self):
                if self.value is not None:
                    string = str(self.value)
                    # remove unneeded digits
                    if isinstance(self.value, float) and string[-2:] == '.0':
                        string = string[:-2]
                    return string
                else:
                    return ''

    class Uno(Interface):
        """
        Handles Uno interfacing with Office program
        """

        class Model(Model):
            """
            Handles usages of PyUno Model
            """

            def __init__(self) -> None:
                # not an error; provided by macro caller
                desktop = XSCRIPTCONTEXT.getDesktop()
                py_uno_model = desktop.getCurrentComponent()
                if not hasattr(py_uno_model, 'Sheets'):
                    raise AttributeError(
                        'Model does not have Sheets. '
                        'This macro needs to be run from a calc workbook.'
                    )
                self.model = py_uno_model

            def __getitem__(self, item: str or int) -> Sheet:
                """
                Gets identified sheet.
                :param item: str or int
                :return: Sheet
                """
                assert isinstance(item, (str, int))
                # try to get appropriate sheet from uno model.
                # If sheet index or name cannot be found, raise a more
                # readable error message than the terribly unhelpful
                # uno message.
                if isinstance(item, int):
                    try:
                        return Office.Uno.Sheet(
                            self.model.Sheets.getByIndex(item))
                    except:  # can't seem to put the actual exception
                        # class here
                        raise IndexError('Could not retrieve sheet at index'
                                         '%s' % repr(item))
                else:
                    try:
                        return Office.Uno.Sheet(
                            self.model.Sheets.getByName(item))
                    except:
                        raise KeyError('Could not retrieve sheet with name %s'
                                       % repr(item))

            def sheet_exists(self, *args: str) -> str:
                """
                Checks each string passed as arg to see if it exists as a
                sheet name. If so, returns it, otherwise, moves to the next
                :param args: strings
                :return: str of first viable sheet name or None if no
                viable name is found
                """
                assert all([isinstance(arg, str) for arg in args])
                for sheet_name in args:
                    try:
                        self.model.Sheets.getByName(sheet_name)
                    except:  # todo: find actual exception class
                        pass
                    else:
                        return sheet_name

            @property
            def sheets(self):
                """
                Generator returning each sheet in Model / Book
                :return: Sheet
                """
                i = 0
                while True:  # loop until break
                    try:
                        yield Office.Uno.Sheet(self.model.Sheets.getByIndex(i))
                    except:
                        break
                    else:
                        i += 1

            @property
            def sheet_names(self):
                """
                Returns tuple of sheet names in model / current book
                :return: tuple
                """
                return self.model.Sheets.ElementNames

        class Sheet(Sheet):
            """
            Handles usage of a workbook sheet
            """

            def __init__(
                    self,
                    uno_sheet,
                    reference_row_index=0,
                    reference_column_index=0
            ) -> None:
                self.i7e_sheet = uno_sheet
                self.reference_row_index = reference_row_index
                self.reference_column_index = reference_column_index

            def get_column_by_index(self, column_index: int) -> Column:
                """
                Gets column of passed index
                :param column_index: int
                :return: Column
                """
                if not isinstance(column_index, int):
                    raise TypeError('Column index must be an int, got %s'
                                    % column_index)
                if column_index < 0:
                    raise ValueError('Passed index must be 0 or greater.')
                return Office.Uno.Column(
                    sheet=self,
                    column_index=column_index,
                    reference_column_index=self.reference_column_index
                )

            def get_row_by_index(self, row_index: int or str) -> Row:
                """
                Gets row of passed index
                :param row_index: int
                :return: Row
                """
                if not isinstance(row_index, int):
                    raise TypeError('row_index must be an int, got %s '
                                    % row_index)
                return Office.Uno.Row(
                    sheet=self,
                    row_index=row_index,
                    reference_row_index=self.reference_row_index
                )

        class Line(Line):
            """
            Contains methods common to both Columns and Rows
            """

            def __len__(self) -> int:
                n = 0
                for each in self:
                    n += 1
                return n

            @property
            def uno_sheet(self):
                """
                Gets uno sheet object
                :return: uno sheet
                """
                return self.sheet.i7e_sheet

        class Column(Line, Column):
            """
            Handles usage of a column within a sheet
            """
            def __init__(
                    self,
                    sheet: Sheet,
                    column_index: int,
                    reference_column_index: int=0):
                super().__init__(
                    sheet=sheet,
                    index=column_index,
                    reference_index=reference_column_index
                )

            def __iter__(self):
                """
                Returns iterable line of cells
                :return: Iterable
                """
                return self.get_iterator(axis='y')

            def get_cell_by_index(self, index: int) -> Cell:
                """
                Gets cell from passed index.
                :param index: int
                :return: Cell
                """
                if not isinstance(index, int):
                    raise TypeError("Passed index must be an int, got %s"
                                    % index)
                return Office.Uno.Cell(
                    self.sheet, (self.index, index))

        class Row(Line, Row):
            """
            Handles usage of a row within a sheet
            """

            def __init__(
                    self,
                    sheet: Sheet,
                    row_index: int,
                    reference_row_index: int=0
            ) -> None:
                super().__init__(
                    sheet=sheet,
                    index=row_index,
                    reference_index=reference_row_index,
                )

            def __iter__(self):
                self.get_iterator(axis='x')

            def get_cell_by_index(self, index: int) -> Cell:
                """
                Gets cell in Row from passed index
                :param index: int
                :return: Cell
                """
                if not isinstance(index, int):
                    raise TypeError('Passed index should be int, got %s'
                                    % index)
                return Office.Uno.Cell(
                    sheet=self.sheet,
                    position=(index, self.index))

        class Cell(Cell):
            """
            Handles usage of an individual cell
            """
            def set_color(self, color):
                """
                Sets cell background color
                :param color: int, list or tuple
                """
                assert isinstance(color, (int, tuple, list))
                if isinstance(color, int):
                    color_int = color
                else:
                    color_int = color[0] * 256 ** 2 + color[1] * 256 + color[2]
                self._source_cell.CellBackColor = color_int

            def get_color(self) -> int:
                return self._source_cell.CellBackColor

            @property
            def _uno_sheet(self):
                """
                Gets uno Sheet obj.
                :return: uno Sheet obj
                """
                return self.sheet.i7e_sheet

            @property
            def _source_cell(self):
                """
                Gets PyUno cell from which values are drawn
                :return:
                """
                return self._uno_sheet.getCellByPosition(*self.position)

            @property
            def value(self) -> int or float or str:
                """
                Gets value of cell.
                :return: str or float
                """
                # get cell value type after formula evaluation has been
                # carried out. This will return the cell value's type
                # even if it is not a formula
                t = self._source_cell.FormulaResultType.value
                if t == 'TEXT':
                    return self._source_cell.getString()
                elif t == 'VALUE':
                    return self._source_cell.getValue()

            @value.setter
            def value(self, new_value: int or float or str) -> None:
                """
                Sets source cell string and number value appropriately for
                a new value.
                This does not handle formulas at the present time.
                :param new_value: int, float, or str
                """
                assert isinstance(new_value, (str, int, float))
                if isinstance(new_value, str):
                    self.string = new_value
                else:
                    self.float = new_value

            @property
            def string(self) -> str:
                """
                Returns string value directly from source cell
                :return: str
                """
                return self._source_cell.getString()

            @string.setter
            def string(self, new_string: str) -> None:
                """
                Sets string value of source cell directly
                :param new_string: str
                """
                assert isinstance(new_string, str)
                self._source_cell.setString(new_string)

            @property
            def float(self) -> float:
                """
                Returns float value directly from source cell 'value'
                :return: float
                """
                return self._source_cell.getValue()

            @float.setter
            def float(self, new_float: int or float) -> None:
                """
                Sets float value of source cell directly
                :param new_float: int or float
                :return: None
                """
                assert isinstance(new_float, (int, float))
                new_value = float(new_float)
                self._source_cell.setValue(new_value)

    @staticmethod
    def get_interface() -> str or None:
        """
        Test for what interface is using this macro, and return the string
        of the appropriate class that should be used.
        :return: str or None if no interface can be determined
        """
        # test for Python Uno
        try:
            XSCRIPTCONTEXT  # if this variable exists, PyUno is being used.
        except NameError:
            pass
        else:
            return 'Uno'

        if xw is not None:
            return 'XW'

        # otherwise, return None / False

    @staticmethod
    def get_interface_class() -> Interface:
        """Gets interface class, ie, Uno or XW"""
        interface = Office.get_interface()  # gets str name of interface class
        if not interface:
            raise ValueError('Should be run as macro using XLWings or PyUno.'
                             'Neither could be detected.')
        return getattr(Office, interface)

    @staticmethod
    def get_model() -> Model:
        return Office.get_model_class()()  # get model class and instantiate

    @staticmethod
    def get_model_class() -> type:
        """Gets appropriate Model class"""
        return Office.get_interface_class().Model

    @staticmethod
    def get_sheet_class() -> type:
        """Gets appropriate Sheet class"""
        return Office.get_interface_class().Sheet

    @staticmethod
    def get_column_class() -> type:
        """Gets appropriate Column class"""
        return Office.get_interface_class().Column

    @staticmethod
    def get_row_class() -> type:
        """Gets appropriate Row class"""
        return Office.get_interface_class().Row

    @staticmethod
    def get_cell_class() -> type:
        """Gets appropriate Cell class"""
        return Office.get_interface_class().Cell


class Color:
    """
    Handles color conversions such as hex -> RGB or RGB -> hex
    """
    color = 0

    def __init__(self, color):
        if isinstance(color, int):
            self.color = color  # store as int
        elif isinstance(color, tuple):
            if len(color) != 3:
                raise ValueError(
                    "Color: RGB tuple should be len 3. got: %s" % color
                )
            if any([(not isinstance(value, int) or value > 255)
                    for value in color]):
                raise ValueError(
                    "Color: each value in color rgb tuple should be an int"
                    "between 0 and 255. Got: %s" % color
                )
            r = color[0]
            g = color[1]
            b = color[2]
            self.color = (r << 16) + (g << 8) + b

        elif isinstance(color, str):
            raise ValueError("Color does not yet support string colors")
        elif isinstance(color, Color):
            self.color = color.color  # color has stopped sounding like a word
        else:
            raise TypeError("Color constructor was passed an "
                            "unknown color identifier: %s" % color)

    @property
    def rgb(self):
        r = (self.color >> 16) & 255
        g = (self.color >> 8) & 255
        b = self.color & 255
        return r, g, b