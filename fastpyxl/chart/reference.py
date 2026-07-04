# Copyright (c) 2010-2024 fastpyxl


from typing import cast

from fastpyxl.utils import (
    get_column_letter,
    range_to_tuple,
    quote_sheetname
)


def _row_col(value, *, lo, hi, name):
    if value is None:
        return None
    value = int(value)
    if value < lo or value > hi:
        raise ValueError(f"{name} must be between {lo} and {hi}")
    return value


class Reference:

    """
    Normalise cell range references
    """

    def __init__(self,
                 worksheet=None,
                 min_col=None,
                 min_row=None,
                 max_col=None,
                 max_row=None,
                 range_string=None,
                 sheet_title=None,
                 ):
        if range_string is not None:
            sheet_title, boundaries = range_to_tuple(range_string)
            min_col, min_row, max_col, max_row = boundaries

        self.worksheet = worksheet
        self._sheet_title = sheet_title
        self.min_col = _row_col(min_col, lo=1, hi=16384, name="min_col")
        self.min_row = _row_col(min_row, lo=1, hi=1000000, name="min_row")
        if max_col is None:
            max_col = min_col
        self.max_col = _row_col(max_col, lo=1, hi=16384, name="max_col")
        if max_row is None:
            max_row = min_row
        self.max_row = _row_col(max_row, lo=1, hi=1000000, name="max_row")
        self.range_string = None

    def __repr__(self):
        return str(self)

    def __str__(self):
        fmt = u"{0}!${1}${2}:${3}${4}"
        if (self.min_col == self.max_col
            and self.min_row == self.max_row):
            fmt = u"{0}!${1}${2}"
        return fmt.format(self.sheetname,
                          get_column_letter(self.min_col), self.min_row,
                          get_column_letter(self.max_col), self.max_row
                          )

    __str__ = __str__

    def __len__(self):
        if self.min_row == self.max_row:
            return 1 + cast(int, self.max_col) - cast(int, self.min_col)
        return 1 + cast(int, self.max_row) - cast(int, self.min_row)

    def __eq__(self, other):
        return str(self) == str(other)

    def _child(self, **kwargs):
        return Reference(
            worksheet=self.worksheet,
            sheet_title=self._sheet_title,
            **kwargs,
        )

    @property
    def rows(self):
        """
        Return all rows in the range
        """
        for row in range(cast(int, self.min_row), cast(int, self.max_row) + 1):
            yield self._child(min_col=self.min_col, min_row=row, max_col=self.max_col, max_row=row)

    @property
    def cols(self):
        """
        Return all columns in the range
        """
        for col in range(cast(int, self.min_col), cast(int, self.max_col) + 1):
            yield self._child(min_col=col, min_row=self.min_row, max_col=col, max_row=self.max_row)

    def pop(self):
        """
        Return and remove the first cell
        """
        cell = "{0}{1}".format(get_column_letter(self.min_col), self.min_row)
        if self.min_row == self.max_row:
            self.min_col = cast(int, self.min_col) + 1
        else:
            self.min_row = cast(int, self.min_row) + 1
        return cell

    @property
    def sheetname(self):
        if self.worksheet is not None:
            return quote_sheetname(self.worksheet.title)
        if self._sheet_title is None:
            raise AttributeError("Reference is missing worksheet context")
        return quote_sheetname(self._sheet_title)
