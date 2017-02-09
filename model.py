try:
    import xlwings as xw
except ImportError:
    xw = None

MAX_CELL_GAP = 10  # max distance between inhabited cells in the workbook
DEFAULT_COLOR = -1

###############################################################################
# BOOK INTERFACE


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
        """
        Gets iterator returning sheets in model
        :return: Iterator
        """
        raise NotImplementedError

    def get_sheet(
            self,
            sheet_name: str,
            row_ref_i: int=0,
            col_ref_i: int=0
    ) -> 'Sheet':
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
        """
        Checks that the passed string(s) exist in the model.
        if so, returns the first passed sheet name to do so.
        :param sheet_name: str
        :return: str
        """
        raise NotImplementedError
        # implemented by office program specific subclasses

    @property
    def sheets(self):
        """
        Returns iterator of sheets in model.
        :return: Iterator
        """
        raise NotImplementedError
        # implemented by office program specific subclasses

    @property
    def sheet_names(self):
        """
        Gets iterable of sheet names
        :return: str iterable
        """
        raise NotImplementedError


class WorkBookComponent:
    """
    Abstract class with common methods for classes that exist within
    a workbook.
    """
    sheet = None

    @property
    def parents(self):
        """
        Gets parents of WorkBookComponent (instances containing this component)
        :return: Generator[WorkBookComponent]
        """
        raise NotImplementedError

    @property
    def instantiated_parents(self):
        """
        Works as 'parents' method, but returns only parents that have
        been instantiated.
        This method was written to facilitate clearing caches when a
        sub-component has been changed
        :return: Generator[WorkBookComponent]
        """
        raise NotImplementedError

    @staticmethod
    def value_cache(getter: callable) -> callable:
        """
        decorator that caches results of the passed method
        :param getter: callable
        :return: callable
        """
        def getter_wrapper(o, *args, **kwargs):
            # if caching is not possible, just call getter
            # to do this, find sheet and check if exclusive_editor is True
            if not o.sheet.exclusive_editor:
                return getter(o, *args, **kwargs)
            # get cache
            try:
                cache = o.__value_cache
            except AttributeError:
                cache = o.__value_cache = {}
            # if caching is possible, but nothing is in cache, set it
            try:
                value = cache[getter]  # try to return cached value
            except KeyError:
                value = cache[getter] = getter(o, *args, **kwargs)
            return value
        return getter_wrapper

    @staticmethod
    def enduring_cache(getter):
        """
        Caches results of method in a cache that is not removed
        when a clear_cache method is called. This is to be used
        for methods whose results will always remain the same
        for the object they have been called upon, even if values
        of that WorkBookObject have changed.
        This method will also cache results regardless of whether
        the sheet interface obj is the exclusive editor of its values,
        since changes made by other means do not affect the returned
        values of the method being cached.
        :param getter: callable
        :return: callable
        """
        def enduring_cache_getter(o, *args, **kwargs):
            # get cache
            try:
                cache = o.__enduring_cache
            except AttributeError:
                cache = o.__enduring_cache = {}
            # if caching is possible, but nothing is in cache, set it
            try:
                value = cache[getter]  # try to return cached value
            except KeyError:
                value = cache[getter] = getter(o, *args, **kwargs)
            return value

        return enduring_cache_getter

    @staticmethod
    def clear_cache(setter):
        """
        decorator method to be used by value setter.
        amends method to clear cache of this WorkBookComponent and
        instantiated parents.
        :param setter: callable
        :return: callable
        """

        def setter_wrapper(o, *args, **kwargs) -> None:
            """
            Function wrapping passed setter method.
            first clears cached values,
            then calls original setter method with values.
            :return: None
            """
            assert isinstance(o, WorkBookComponent)
            # for each WorkBookComponent modified, clear its cache
            modified_components = \
                [o] + [parent for parent in o.instantiated_parents]
            for component in modified_components:
                try:  # try to find cache.
                    cache = component.__value_cache
                except AttributeError:
                    pass
                else:  # if cell has a cache, clear it.
                    assert isinstance(cache, dict)
                    cache.clear()
            setter(o, *args, **kwargs)  # call original value setter method

        return setter_wrapper


class Sheet(WorkBookComponent):
    """
    Abstract class for sheet handling cell values of a passed sheet.
    Inherited from by PyUno.Sheet and XW.Sheet.
    """
    i7e_sheet = None  # interface sheet obj. ie; com.sun.star...Sheet
    _reference_row_index = 0
    _reference_column_index = 0
    _content_row_start = None  # default start index of table_row iter
    _table_col_start = None  # default start index of table_col iter
    _snapshot = None
    exclusive_editor = False  # true if no concurrent editing will occur

    _all_sheets = {}  # all sheets so far instantiated, by repr string

    def __init__(
            self,
            i7e_sheet,  # used by subclasses
            reference_row_index=0,  # used by subclasses
            reference_column_index=0  # used by subclasses
    ) -> None:
        self.i7e_sheet = i7e_sheet
        self.reference_column_index = reference_column_index
        self.reference_row_index = reference_row_index
        self.sheet = self  # returns self, as required by WorkBookComponent
        assert repr(self) not in Sheet._all_sheets
        Sheet._all_sheets[repr(self)] = self

    @staticmethod
    def factory(i7e_sheet, ref_row_index=0, ref_col_index=0) -> 'Sheet':
        """
        Factory method return Sheet.
        This method does not create duplicate sheets,
        instead, it returns a reference to a pre-existing sheet
        if one exists.
        :return: Sheet
        """
        sheet_class = Office.get_sheet_class()
        key = sheet_class.key(i7e_sheet)
        try:
            return Sheet._all_sheets[key]
        except KeyError:
            return sheet_class(
                i7e_sheet=i7e_sheet,
                reference_column_index=ref_col_index,
                reference_row_index=ref_row_index
            )

    @staticmethod
    def key(i7e_sheet):
        return i7e_sheet

    def get_column(
            self,
            column_identifier: int or float or str
    ) -> 'Column':
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

    def get_column_by_index(self, column_index: int) -> 'Column':
        """
        Gets column in Sheet by passed index.
        Implemented by sub-classes.
        :param column_index: int
        :return: Column
        """
        return Column.factory(
            sheet=self,
            index=column_index,
            reference_index=self.reference_column_index
        )

    def get_column_by_name(self, column_name: int or float or str) -> 'Column':
        """
        Gets column from a passed reference value which is compared
        to each cell value in the reference row.
        This function will return the first column whose name matches
        the passed value.
        :return: Office.Column
        """
        i = self.get_column_index_from_name(column_name)
        return self.get_column_by_index(i) if i is not None else None

    def get_column_index_from_name(
        self,
        column_name: int or float or str
    ) -> int or None:
        """
        Gets column index from name
        :param column_name: int, float, or str
        :return: int or None
        """
        # todo: cache this
        for x, cell in enumerate(self.reference_row):
            if cell.value == column_name:
                return x

    def get_row(self, row_identifier: int or float or str) -> 'Row':
        """
        Gets row by name if identifier is str, otherwise by index
        :param row_identifier: int, float, or str
        :return: Office.Row
        """
        if isinstance(row_identifier, str):
            return self.get_row_by_name(row_identifier)
        else:
            return self.get_row_by_index(row_identifier)

    def get_row_by_index(self, row_index: int or str) -> 'Row':
        """
        Gets row by passed index.
        Implemented by subclasses.
        :param row_index: int
        :return: Row
        """
        return Row.factory(
            sheet=self,
            index=row_index,
            reference_index=self.reference_row_index
        )

    def get_row_by_name(self, row_name: int or str or float) -> 'Row':
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
        for y, cell in enumerate(self.reference_column):
            if cell.value == row_name:
                return y

    def get_cell(self, cell_identifier, **kwargs) -> 'Cell':
        """
        Gets cell from Sheet.
        May be identified as;
        example_sheet.get_cell((x, y))
        example_sheet.get_cell((row_name, column_name))
        example_sheet.get_cell((row_name, y_int))
        Assumes that a passed number is a row or column index.
        If a number is to be passed as a row or sheet name,
        pass as a kwarg x_identifier_type='name' or
        y_identifier_type='name'.
        Also valid: x_identifier_type='index'
        :param cell_identifier: tuple
        :param kwargs: strings
        :return:
        """
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
                x = self.get_column_index_from_name(x_identifier)
            else:
                assert isinstance(x_identifier, int)  # sanity check
                x = x_identifier
            if (y_identifier_type == 'name' or
                y_identifier_type is None and isinstance(
                    y_identifier, str)):
                y = self.get_row_index_from_name(y_identifier)
            else:
                assert isinstance(y_identifier, int)  # sanity check
                y = y_identifier
            # We now have x and y indices for the cell
            # now we call the cell factory method, which returns a cell
            # of the correct type. The Cell factory method will not
            # create duplicate cells, instead it will return a
            # reference to the pre-existing cell in that sheet+position
            return Cell.factory(sheet=self, position=(x, y))

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
    def reference_row(self) -> 'Row':
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
    def reference_column(self) -> 'Column':
        """
        Gets reference column
        :return: Office.Column
        """
        return self.get_column(self.reference_column_index)

    @property
    def columns(self) -> 'LineSeries':
        """
        Gets iterator of columns in sheet
        :return: Iterator[Column]
        """
        return LineSeries(reference_line=self.reference_row)

    @property
    def table_columns(self) -> 'LineSeries':
        """
        Gets iterator of columns in sheet's table contents
        :return: Iterator[Column]
        """
        return LineSeries(
            reference_line=self.reference_column,
            start_index=self.table_col_start_i
        )

    @property
    def table_col_start_i(self) -> int:
        """
        Gets index of first column in sheet's table.
        By default, this is the first column after the reference column.
        This value may also be explicitly set by passing an int to this
        property setter.
        :return: int
        """
        if self._table_col_start is not None:
            return self._table_col_start
        else:
            return self._reference_column_index + 1

    @table_col_start_i.setter
    def table_col_start_i(self, new_i: int) -> None:
        """
        Sets first column which is included in table_columns.
        :param new_i: int
        :return: None
        """
        if new_i is not None or not isinstance(new_i, int):
            raise TypeError(
                'New index should be an integer, or None. Got: %s'
                % repr(new_i)
            )
        self._table_col_start = new_i

    @property
    def rows(self) -> 'LineSeries':
        """
        Gets iterator of rows in Sheet.
        :return: Iterator[Row]
        """
        return LineSeries(reference_line=self.reference_column)

    @property
    def table_rows(self) -> 'LineSeries':
        """
        Gets iterator of table rows in Sheet
        :return: Iterator[Row]
        """
        return LineSeries(
            reference_line=self.reference_column,
            start_index=self.table_row_start_i
        )

    @property
    def table_row_start_i(self) -> int:
        """
        Gets index of first row in sheet's table.
        By default, this is the first row after the reference row.
        This value may also be explicitly set by passing an int to this
        property setter.
        :return: int
        """
        if self._content_row_start is not None:
            return self._content_row_start
        else:
            return self.reference_row_index + 1

    @table_row_start_i.setter
    def table_row_start_i(self, new_i: int) -> None:
        """
        Sets first row which is included in table_columns.
        :param new_i: int
        :return: None
        """
        if new_i is not None or not isinstance(new_i, int):
            raise TypeError(
                'New index should be an integer, or None. Got: %s'
                % repr(new_i)
            )
        self._content_row_start = new_i

    def take_snapshot(
            self,
            width: int=None,
            height: int=None,
            frozen_size: bool=False
    ):
        """
        Gets snapshot of sheet as it currently exists, and uses that
        for all reading and writing of values.
        The snapshot will by default include all values within the
        width of the reference row, and within the height of the
        reference column. Values outside that area may be overridden
        if the range is expanded.
        The snapshot can be prevented from growing by setting
        frozen_size to True.
        If a snapshot has already been taken, it is replaced, along
        with any changes made to it.
        :param width: int
        :param height: int
        :param frozen_size: bool
        :return: None
        """
        if height is None:
            height = len(self.reference_column)
        if width is None:
            width = len(self.reference_row)
        if height == 0 and width > 0:
            height = 1  # if width > 0, then a reference row was found
        if width == 0 and height > 0:
            width = 1  # if height > 0, then a reference col was found
        self._snapshot = self.Snapshot(self, width, height, frozen_size)

    def write_snapshot(self):
        """
        Writes values in snapshot to memory in one single write
        (much, much faster than writing each cell individually)
        :return: None
        """
        self._snapshot.write()

    def discard_snapshot(self):
        """
        Throws out snapshot along with all changes made to it
        :return: None
        """
        self._snapshot = None

    @property
    def snapshot(self) -> None or 'Snapshot':
        return self._snapshot

    @property
    def screen_updating(self) -> bool or None:
        """
        Gets bool of whether screen updating occurs on sheet.
        does nothing here, may be overridden by subclasses.
        returns None if this sheet type does not support this method.
        :return: bool
        """
        return

    @screen_updating.setter
    def screen_updating(self, new_bool: bool) -> None:
        """
        Sets whether sheet refreshes its display after update of
        values, etc.
        Does nothing here, subclasses may implement this method.
        :param new_bool: bool
        :return: None
        """
        return

    @property
    def parents(self):
        return  # nothing to yield or return

    @property
    def instantiated_parents(self):
        return  # nothing to yield or return

    def __str__(self) -> str:
        raise NotImplementedError

    def __repr__(self) -> str:
        raise NotImplementedError

    class Snapshot:
        """
        Class storing a sheet entirely in memory, so that individual
        reads and writes to a sheet can be grouped together.
        Snapshot is sub-classed within XW and Uno.
        """

        def __init__(
                self,
                sheet: 'Sheet',
                width: int,
                height: int,
                frozen_size: bool
        ) -> None:
            self._sheet = sheet
            self._height = height
            self._width = width
            self.frozen_size = frozen_size
            self._values = self._get_values()

        def _get_values(self) -> list:
            """
            Gets list of lists containing values of cells within snapshot
            :return: list[list[str, float or None]]
            """
            raise NotImplementedError

        def _grow(self, new_width: int=None, new_height: int=None) -> None:
            """
            Grows snapshot to new passed size.
            If passed width or height is smaller than the present size,
            the snapshot will stay the same size, it will not shrink.
            :param new_width: int
            :param new_height: int
            :return: None
            """
            if new_height < self._height:
                new_height = self._height
            if new_width < self._width:
                new_width = self._width
            height_difference = new_height - self._height
            width_difference = new_width - self._width
            # first add rows
            if height_difference:
                self._values += [[None] * new_width] * height_difference
            # then columns
            if width_difference:
                for i in range(0, self._height):  # for each pre-existing row
                    self._values[i] += [None] * width_difference
            self._width = new_width
            self._height = new_height

        def get_value(self, x: int, y: int) -> str or float or int or None:
            """"
            Gets value of cell at x, y.
            Throws IndexError if x or y is outside range of snapshot.
            :param x: int
            :param y: int
            :return str, float, int or None
            """
            try:
                row = self._values[y]
            except IndexError:
                raise IndexError('Snapshot:get_value: %s is outside height of '
                                 'snapshot. (height:%s)' %
                                 (x, self._height))
            else:
                try:
                    return row[x]  # return value
                except IndexError:
                    raise IndexError(
                        'Snapshot:get_value: %s is outside width of '
                        'snapshot. (width:%s)' %
                        (x, self._width))

        def set_value(self, x: int, y: int, value) -> None:
            """
            Sets value of cell at x, y. (to be committed to cell when
            snapshot is written.)
            Throws IndexError if x or y is outside range of snapshot
            :param x: int
            :param y: int
            :param value: Any
            :return: None
            """
            if not isinstance(value, (int, float, str, None)):
                raise TypeError(
                    'Snapshot:set_value: Passed value should be an int, float'
                    ' , str, or None. Got: %s' % repr(value))
            if x >= self._width or y >= self._height:
                if self.frozen_size:
                    raise IndexError(
                        'Snapshot:set_value: (%s, %s) is outside range of '
                        'snapshot (size: (%s, %s)' %
                        (x, y, self._width, self._height)
                    )
                self._grow(x + 1, y + 1)
            self._values[y][x] = value

        def write(self):
            """
            Commits values in snapshot to sheet in a single bulk write
            :return: None
            """
            raise NotImplementedError


class LineSeries:
    """Class storing collection of Line, Column, or Row objects"""

    COLUMNS_STR = 'columns'
    ROWS_STR = 'rows'

    def __init__(
            self,
            reference_line: 'Line',
            start_index: int=0,  # iterator start index
            end_index: int=None,  # iterator end index (exclusive)
    ) -> None:
        if not isinstance(reference_line, Line):
            raise TypeError('Expected LineSeries to be passed reference line.'
                            'Got instead: %s' % repr(reference_line))
        self.reference_line = reference_line
        self.start_index = start_index
        self.end_index = end_index

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
        # find whether this is a series of rows or columns
        if self.end_index is not None:
            ref_cells = self.reference_line[self.start_index:self.end_index]
        else:
            ref_cells = self.reference_line[self.start_index:]
        for cell in ref_cells:
            if self._contents_type == LineSeries.COLUMNS_STR:
                yield cell.column
            if self._contents_type == LineSeries.ROWS_STR:
                yield cell.rows

    def __len__(self) -> int:
        """
        Returns size of LineSeries
        :return: int
        """
        count = self.start_index
        try:  # it feels like there should be a better way to do this
            assert self.end_index is None or isinstance(self.end_index, int)
            # while there is a next cell, and it is inside range
            while self.__iter__().__next__() and \
                    (self.end_index is None or count + 1 < self.end_index):
                count += 1  # increment count
        except StopIteration:
            return count

    def get_by_name(self, name: int or float or str) -> 'Line':
        """
        Gets line from passed line name.
        Returns None if no line of that name is found.
        :param name: int, float or str
        :return: Line or None
        """
        for cell in self.reference_line:
            if cell.value == name:
                if self._contents_type == LineSeries.COLUMNS_STR:
                    return cell.column
                elif self._contents_type == LineSeries.ROWS_STR:
                    return cell.row

    def get_by_index(self, index: int) -> 'Line':
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
        """
        Gets the str name of line series; either 'columns' or 'rows'
        :return: str
        """
        if isinstance(self.reference_line, Row):
            return 'columns'
        elif isinstance(self.reference_line, Column):
            return 'rows'
        else:
            raise TypeError('Expected reference_line to be Row or'
                            'Column. Got: %s' % repr(self.reference_line))


class Line(WorkBookComponent):
    """
    Abstract class for a line of cells.
    Sub-classed by both Row and Column
    """
    _all = {}  # overwritten in sub-classes.

    def __init__(
        self,
        sheet: Sheet,
        index: int,
        reference_index: int,
    ) -> None:
        # check if an instance already exists in dict of all Row/Col.
        # It shouldn't.
        assert (sheet, index) not in self._all
        self._all[(sheet, index)] = self  # add self to _all dict
        if not isinstance(sheet, Sheet):
            raise TypeError(
                'Expected subclass of Sheet. Instead got %s'
                % repr(sheet))
        if not isinstance(index, int):
            raise TypeError(
                'Expected line index to be an int. '
                'Instead got %s' % repr(index))
        if index < 0:
            raise ValueError('Index must be 0 or greater. Got: %s' % index)
        if not isinstance(reference_index, int):
            raise TypeError(
                'Expected reference name index to be an int, '
                'Instead got %s' % repr(reference_index))
        if reference_index < 0:
            raise ValueError('ref index must be 0 or greater. Got: %s' % index)
        self.sheet = sheet
        self.index = index
        self.reference_index = index

    def __getitem__(self, cell_identifier):  # returns Cell or Generator
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
        assert isinstance(cell_identifier, (int, float, str, slice)), \
            'Expected cell_identifier to be number, str, or slice got %s ' \
            % cell_identifier
        if isinstance(cell_identifier, slice):
            return self.slice(cell_identifier)
        if isinstance(cell_identifier, int):
            return self.get_cell_by_index(cell_identifier)
        else:
            return self.get_cell_by_reference(cell_identifier)

    def __iter__(self):
        raise NotImplementedError
        # implemented by office program specific subclasses

    @WorkBookComponent.value_cache
    def __len__(self) -> int:
        count = 0
        for _ in self:
            count += 1
        return count

    @classmethod
    def exists(cls, sheet: Sheet, index: int):
        """
        Returns bool of whether a line of sheet+index already exists
        :param sheet: Sheet
        :param index: int
        :return: bool
        """
        return sheet, index in cls._all

    def slice(self, s: slice):  # return cell generator
        """
        Returns generator returning items in slice
        :param s: slice
        :return: Generator[Cell]
        """
        if s.start is None:
            start = 0
        else:
            start = s.start
        if s.stop is None:
            stop = len(self)
        else:
            stop = s.stop
        if s.step is not None:
            rng = range(start, stop, s.step)
        else:
            rng = range(start, stop)
        for i in rng:
            if i in range(len(self)):
                yield self[i]

    def get_cell_by_index(self, index: int) -> 'Cell':
        """
        Gets cell in Line at passed index.
        :param index: int
        :return: Cell
        """
        raise NotImplementedError
        # implemented by Row and Column in office program
        # specific subclasses

    def get_cell_by_reference(self, reference: str or float or int) -> 'Cell':
        for i, cell in enumerate(self._reference_line):
            if cell.value == reference:
                return self.get_cell_by_index(i)

    def get_iterator(self, axis: str) -> 'CellLine':
        assert axis == 'x' or axis == 'y'
        return CellLine(self.sheet, axis, self.index)

    def clear(self, include_header: bool = False) -> None:
        """
        Clears line of cells.
        If Include header is True; clears cell data in cells
        preceding and including header row.
        :param include_header: bool
        :return: None
        """
        [cell.clear() for i, cell in enumerate(self)
         if i > self.name_cell_index or include_header]

    @property
    def _reference_line(self) -> 'Line':
        """
        Gets reference line which is parallel to this Line and
        which is used to look up names of cells in this Line.
        Implemented in sub-classes.
        :return: Line
        """
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
    def name(self) -> int or float or str or None:
        """
        Returns name of line, which is the value stored in the
        line's header cell, located in the sheet's reference
        row or column.
        :return: int, float, str or None
        """
        return self[self.name_cell_index].value

    def to_dict(self) -> dict:
        """
        Returns line values as dictionary, with cell values as values,
        and corresponding reference row values as keys.
        :return: dict
        """
        return {line_cell.value: ref_cell.value for
                line_cell, ref_cell in zip(self, self._reference_line)}

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

    @property
    def parents(self):
        yield self.sheet

    @property
    def instantiated_parents(self):
        yield self.sheet


class Column(Line):
    """
    Abstract Column class, extended by Office.XW.Column and Office.Uno.Column
    """
    _all = {}  # dict storing all unique sheet+index possibilities

    @staticmethod
    def factory(sheet: 'Sheet', index: int, reference_index: int) -> 'Column':
        """
        Factory returning a Row belonging to sheet, at passed index,
        with cell name/keys determined by reference_index.
        :param sheet: Sheet
        :param index: int
        :param reference_index: int
        :return: Column
        """
        # first try to find a pre-existing row with that sheet + index
        column_key = sheet, index  # sheet instance is expected to be unique
        try:
            column = Column._all[column_key]
        except KeyError:
            column = Office.get_column_class()(sheet, index, reference_index)
        return column

    def get_cell_by_index(self, index: int) -> 'Cell':
        """
        Gets cell in Column at passed index.
        :param index: int
        :return: Cell
        """
        if not isinstance(index, int):
            raise TypeError(
                'get_cell_by_index: passed non-int index: %s' % index
            )
        if index < 0:
            raise ValueError(
                'get_cell_by_index: passed index is < 0: %s' % index
            )
        return self.sheet.get_cell((self.index, index))

    @property
    def _reference_line(self) -> 'Line':
        return self.reference_column

    @property
    def reference_column(self) -> 'Column':
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
    _all = {}  # dict storing all unique sheet+index possibilities

    @staticmethod
    def factory(sheet: 'Sheet', index: int, reference_index: int) -> 'Row':
        """
        Factory returning a Row belonging to sheet, at passed index,
        with cell name/keys determined by reference_index.
        :param sheet: Sheet
        :param index: int
        :param reference_index: int
        :return: Row
        """
        # first try to find a pre-existing row with that sheet + index
        row_key = sheet, index
        try:
            row = Row._all[row_key]
        except KeyError:
            row = Office.get_row_class()(sheet, index, reference_index)
        return row

    def get_cell_by_index(self, index: int) -> 'Cell':
        """
        Gets cell in Row at passed index.
        :param index: int
        :return: Cell
        """
        if not isinstance(index, int):
            raise TypeError(
                'get_cell_by_index: passed non-int index: %s' % index
            )
        if index < 0:
            raise ValueError(
                'get_cell_by_index: passed index is < 0: %s' % index
            )
        return self.sheet.get_cell((index, self.index))

    @property
    def _reference_line(self) -> 'Line':
        return self.reference_row

    @property
    def reference_row(self) -> 'Row':
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


class Cell(WorkBookComponent):
    """ Class handling usage of a single cell in office worksheet """
    _all_cells = {}  # dict of all created cells, prevents duplication

    def __init__(
            self,
            sheet: Sheet,
            position: tuple
    ) -> None:
        assert len(position) == 2
        assert all([isinstance(item, int) for item in position])
        self.position = tuple(position)
        self.sheet = sheet

    @staticmethod
    def factory(sheet: Sheet, position: tuple) -> 'Cell':
        """
        Gets cell belonging to sheet, with passed position.
        If a cell with the same sheet + position already exists,
        returns that cell instead.
        :param sheet: Sheet
        :param position: tuple[int, int]
        :return: Cell
        """
        # first try to return an existing cell
        cell_key = repr(sheet), position
        try:
            cell = Cell._all_cells[cell_key]
        except KeyError:
            # if cell does not exist, create it
            Cell._all_cells[cell_key] = cell = \
                Office.get_cell_class()(sheet, position)
        return cell

    def set_color(self, color: int or list or tuple or Color) -> None:
        """
        Sets color in cell to that passed as
        integer (as in a color hex code),
        list or tuple (containing r,g,b values)
        or a Color obj.
        :param color: int, list, tuple, or Color
        :return: None
        """
        raise NotImplementedError

    def get_color(self) -> int:
        """
        Gets color contained in cell as an integer.
        :return: int
        """
        raise NotImplementedError

    def remove_whitespace(self) -> None:
        """
        Edits cell value to remove tab, linefeed, return, form-feed,
        and vertical tab characters.
        :return: None
        """
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
        """
        Gets value contained in sheet cell.
        Does not return only number values.
        :return: int, float, str, or None
        """
        raise NotImplementedError

    @value.setter
    def value(self, new_value: str or int or float or None) -> None:
        """
        Sets cell value.
        :param new_value: str, int, float, or None
        :return: None
        """
        raise NotImplementedError

    @property
    def string(self) -> str:
        """
        Gets string value of cell.
        If stored value is str, returns str as is. Otherwise gets
        string of stored value, or empty string if None.
        :return: str
        """
        raise NotImplementedError

    @string.setter
    def string(self, new_string: str) -> None:
        """
        Sets string value of cell
        :param new_string: str
        :return: None
        """
        raise NotImplementedError

    @property
    def float(self):
        """
        Gets float value of cell.
        If cell contains a string or None, returns 0.
        :return: float
        """
        raise NotImplementedError

    @float.setter
    def float(self, new_float: int or float) -> None:
        """
        Sets float value of cell.
        :param new_float: int or float
        :return:
        """
        raise NotImplementedError

    @property
    def x(self) -> int:
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

    @property
    def parents(self):
        """
        Returns generator iterating over parent instances.
        In the case of Cell this means row, column, sheet instances
        :return: Generator[WorkBookComponent]
        """
        yield self.row
        yield self.column
        yield self.sheet

    @property
    def instantiated_parents(self):
        """
        Returns parents of this cell that have been instantiated.
        This method was written for the purpose of clearing parent
        caches.
        :return: Generator[WorkBookComponent]
        """
        # yield row
        if Row.exists(self.sheet, self.y):
            yield self.row
        if Column.exists(self.sheet, self.x):
            yield self.row
        yield self.sheet

    def __repr__(self) -> str:
        """
        Gets cell repr with sheet, position, and value
        :return: str
        """
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

    def __next__(self) -> 'Cell':
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
            """
            XW data model, extends Model superclass
            """
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
                    return Sheet.factory(
                        self.active_app.books[book_name].sheets[sheet_name]
                    )
                else:
                    # otherwise just look everywhere
                    for sheet in self._xw_sheets:
                        if sheet.name == item:
                            return Sheet.factory(
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
                    yield Sheet.factory(xw_sheet)

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
            """
            XW Sheet
            """
            def __init__(
                    self,
                    i7e_sheet,
                    reference_row_index=0,
                    reference_column_index=0
            ) -> None:
                super().__init__(
                    i7e_sheet=i7e_sheet,
                    reference_column_index=reference_column_index,
                    reference_row_index=reference_row_index
                )

            @staticmethod
            def key(i7e_sheet):
                return 'Sheet[%s::%s]' % (
                    i7e_sheet.name,
                    i7e_sheet.book.fullname,
                )

            @property
            def screen_updating(self) -> None:
                return self.i7e_sheet.book.app.screen_updating

            @screen_updating.setter
            def screen_updating(self, new_bool: bool) -> None:
                self.i7e_sheet.book.app.screen_updating = new_bool

            @Sheet.enduring_cache
            def __str__(self) -> str:
                return 'Sheet[%s::%s]' % (
                    self.i7e_sheet.name,
                    self.i7e_sheet.book.name,
                )

            @Sheet.enduring_cache  # this Sheet's repr should not change
            def __repr__(self) -> str:
                return 'Sheet[%s::%s]' % (
                    self.i7e_sheet.name,
                    self.i7e_sheet.book.fullname,
                )

            class Snapshot(Sheet.Snapshot):
                """
                Snapshot of an XW sheet.
                """

                def _get_values(self) -> list:
                    assert isinstance(self._width, int)
                    assert isinstance(self._height, int)
                    assert (self._width == 0) == (self._height == 0)
                    if self._width == 0 and self._height == 0:
                        values = [[]]
                    elif self._width == 1 and self._height == 1:
                        values = [[self._sheet.i7e_sheet.range('A1').value]]
                    elif self._height == 1:
                        values = [self._sheet.i7e_sheet.range(
                            'A1', (1, self._width)
                        ).value]
                    elif self._width == 1:
                        values = [[value] for value in
                                  self._sheet.i7e_sheet.range(
                                  'A1', (self._height, 1)
                                  )]
                    elif self._width > 1 and self._height > 1:
                        values = self._sheet.i7e_sheet.range(
                            'A1',  # start cell
                            (self._height, self._width)  # end cell
                        ).value
                    else:
                        raise ValueError(
                            'Snapshot:_get_values: width, height were %s, %s' %
                            self._width, self._height)
                    assert values is not None  # just sanity checking
                    return values

                def write(self):
                    self._sheet.i7e_sheet.range('A1').value = self._values

        class Line(Line):
            """
            XW Line
            """
            pass  # no methods defined here at this time.
            # keeping for MR0 purposes

        class Column(Line, Column):
            """
            XW Column
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
                return self.get_iterator(axis='y')

        class Row(Line, Row):
            """
            XW Row
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
                return self.get_iterator(axis='x')

        class Cell(Cell):
            """
            XW Cell
            """

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
            @Cell.value_cache
            def value(self) -> int or float or str or None:
                if self.sheet.snapshot:  # if sheet has a snapshot:
                    try:  # try to get value from snapshot
                        return self.sheet.snapshot.get_value(self.x, self.y)
                    except IndexError:
                        pass  # if it does not contain this cell x,y: get range
                return self._range.value

            @value.setter
            @Cell.clear_cache
            def value(self, new_v) -> None:
                if self.sheet.snapshot:  # if sheet has a snapshot
                    try:
                        self.sheet.snapshot.set_value(self.x, self.y, new_v)
                        return
                    except IndexError:
                        pass  # if value outside snapshot bounds, set normally
                self._range.value = new_v

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
                # noinspection PyUnresolvedReferences
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
            Handles usage of a single Uno workbook sheet
            """

            def __init__(
                    self,
                    i7e_sheet,
                    reference_row_index=0,
                    reference_column_index=0
            ) -> None:
                super().__init__(
                    i7e_sheet=i7e_sheet,
                    reference_row_index=reference_row_index,
                    reference_column_index=reference_column_index
                )

        class Line(Line):
            pass  # no methods defined here anymore,
            # keeping this in place for MRO purposes

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
                return self.get_iterator(axis='x')

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
            @Cell.value_cache
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
            @Cell.clear_cache
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
            @Cell.value_cache
            def string(self) -> str:
                """
                Returns string value directly from source cell
                :return: str
                """
                return self._source_cell.getString()

            @string.setter
            @Cell.clear_cache
            def string(self, new_string: str) -> None:
                """
                Sets string value of source cell directly
                :param new_string: str
                """
                assert isinstance(new_string, str)
                self._source_cell.setString(new_string)

            @property
            @Cell.value_cache
            def float(self) -> float:
                """
                Returns float value directly from source cell 'value'
                :return: float
                """
                return self._source_cell.getValue()

            @float.setter
            @Cell.clear_cache
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
            # if this variable exists, PyUno is being used as a macro.
            XSCRIPTCONTEXT
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
