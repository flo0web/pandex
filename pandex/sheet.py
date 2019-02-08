import copy
from enum import Enum


class Side(Enum):
    RIGHT = 1
    BOTTOM = 2


class Cursor:
    def __init__(self, row=0, col=0):
        self._row = row
        self._col = col

    @property
    def row(self):
        return self._row

    @row.setter
    def row(self, value):
        self._row = value

    @property
    def col(self):
        return self._col

    @col.setter
    def col(self, value):
        self._col = value

    def __str__(self):
        return 'Cursor(row={row}, col={col})'.format(row=self._row, col=self._col)


class Group(object):
    def __init__(self, sheet, cursor: Cursor):
        self._sheet = sheet

        self._cursor_start = cursor
        self._cursor_end = copy.deepcopy(cursor)

    def add(self, obj, side: Side = Side.RIGHT, margin_rows: int = 0, margin_cols: int = 0):
        if side == Side.RIGHT:
            self._add_right(obj, margin_rows, margin_cols)
        elif side == Side.BOTTOM:
            self._add_bottom(obj, margin_rows, margin_cols)
        else:
            raise ValueError('Side can be only right or bottom, you use: ' % side)

    def get_cursors(self):
        return self._cursor_start, self._cursor_end

    def _add_right(self, obj, margin_rows, margin_cols):
        obj_cursor = Cursor(
            row=self._cursor_start.row + margin_rows,
            col=self._cursor_end.col + margin_cols
        )

        self._write_object(obj_cursor, obj)

    def _add_bottom(self, obj, margin_rows, margin_cols):
        obj_cursor = Cursor(
            row=self._cursor_end.row + margin_rows,
            col=self._cursor_start.col + margin_cols
        )

        self._write_object(obj_cursor, obj)

    def _write_object(self, obj_cursor, element):
        element.write(self._sheet.workbook, self._sheet.worksheet, obj_cursor)
        self._update_cursor_end(obj_cursor)

    def _update_cursor_end(self, obj_cursor):
        self._cursor_end = obj_cursor
        self._sheet.update_cursor(self._cursor_end)


class Sheet(object):
    def __init__(self, workbook, name):
        self.workbook = workbook
        self.worksheet = workbook.add_worksheet(name)

        self._cursor = Cursor()

    def create_shape(self, side: Side = Side.RIGHT, margin_rows: int = 0, margin_cols: int = 0):  # table = XLTable
        if side == Side.RIGHT:
            return self._add_right(margin_rows, margin_cols)
        elif side == Side.BOTTOM:
            return self._add_bottom(margin_rows, margin_cols)
        else:
            raise ValueError('Side can be only right or bottom')

    def update_cursor(self, cursor: Cursor):
        self._cursor.row = max(self._cursor.row, cursor.row)
        self._cursor.col = max(self._cursor.col, cursor.col)

    @property
    def cursor(self):
        return self._cursor

    def _add_right(self, margin_rows, margin_cols):
        cursor = Cursor(
            row=0 + margin_rows,
            col=self._cursor.col + margin_cols
        )

        return Group(self, cursor)

    def _add_bottom(self, margin_rows, margin_cols):
        cursor = Cursor(
            row=self._cursor.row + margin_rows,
            col=0 + margin_cols
        )

        return Group(self, cursor)
