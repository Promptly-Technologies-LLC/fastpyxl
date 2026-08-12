# Copyright (c) 2010-2024 fastpyxl

"""Indexed, on-demand worksheet reads (row offset index + seekable XML)."""

from __future__ import annotations

from fastpyxl.cell.read_only import EMPTY_CELL, ReadOnlyCell
from fastpyxl.reader.part_cache import DecompressedPart
from fastpyxl.reader.xml_index import build_row_index, iter_row_spans
from fastpyxl.xml.constants import SHEET_MAIN_NS
from fastpyxl.xml.functions import fromstring

from ._read_only import ReadOnlyWorksheet
from ._reader import WorkSheetParser


_ROW_WRAP_PREFIX = f'<sheetData xmlns="{SHEET_MAIN_NS}">'.encode("ascii")
_ROW_WRAP_SUFFIX = b"</sheetData>"
# Excel emits double quotes; accept single-quoted third-party XML too.
_SHARED_MARKERS = (b't="shared"', b"t='shared'")


def _has_shared_formula(data) -> bool:
    return any(marker in data for marker in _SHARED_MARKERS)


class IndexedWorksheet(ReadOnlyWorksheet):
    """Read-only worksheet with a row-level offset index into decompressed XML."""

    def __init__(
        self,
        parent_workbook,
        title,
        worksheet_path,
        shared_strings,
        part=None,
        *,
        rich_text=False,
    ):
        self.parent = parent_workbook
        self.title = title
        self.sheet_state = "visible"
        self._current_row = None
        self._min_row = 1
        self._max_row = None
        self._min_col = None
        self._max_col = None
        self._min_column = 1
        self._max_column = None
        self._bounds_dirty = False
        self._worksheet_path = worksheet_path
        self._shared_strings = shared_strings
        self._rich_text = rich_text
        self._closed = False

        if part is None:
            part = DecompressedPart.from_archive(parent_workbook._archive, worksheet_path)
        self._part = part

        # Scan via bytes_view so spooled parts are mmap'd rather than re-read
        # into a second full Python bytes object.
        with self._part.bytes_view() as data:
            self._row_index = build_row_index(data)
            self._shared_formulae = self._collect_shared_formulae(data)
        self._load_dimensions()

        from fastpyxl.workbook.defined_name import DefinedNameDict

        self.defined_names = DefinedNameDict()

    def _load_dimensions(self) -> None:
        with self._part.open() as src:
            parser = WorkSheetParser(src, [])
            dimensions = parser.parse_dimensions()
        if dimensions is not None:
            self._min_column, self._min_row, self._max_column, self._max_row = dimensions
        elif self._row_index:
            self._min_row = min(self._row_index)
            self._max_row = max(self._row_index)

    def _collect_shared_formulae(self, data) -> dict:
        """Parse rows that declare shared formulas so dependents can translate."""
        shared: dict = {}
        if not _has_shared_formula(data):
            return shared
        parser = self._make_parser(None)
        parser.shared_formulae = shared
        for _row_num, start, end in iter_row_spans(data):
            chunk = data[start:end]
            if not _has_shared_formula(chunk):
                continue
            # Only rows that may introduce a master (non-empty <f> body) need parse.
            if b"</f>" not in chunk:
                continue
            # mmap slices are bytes; wrap for ElementTree.
            if not isinstance(chunk, (bytes, bytearray)):
                chunk = bytes(chunk)
            row_elem = fromstring(_ROW_WRAP_PREFIX + chunk + _ROW_WRAP_SUFFIX)[0]
            parser.parse_row(row_elem)
        return shared

    def _make_parser(self, source) -> WorkSheetParser:
        return WorkSheetParser(
            source,
            self._shared_strings,
            data_only=self.parent.data_only,
            epoch=self.parent.epoch,
            date_formats=self.parent._date_formats,
            timedelta_formats=self.parent._timedelta_formats,
            rich_text=self._rich_text,
            keep_formula_cache=self.parent.keep_formula_cache,
        )

    def _get_source(self):
        self._ensure_open()
        return self._part.open()

    def _parse_indexed_row(self, row: int):
        span = self._row_index.get(row)
        if span is None:
            return row, []
        start, end = span
        chunk = self._part.read_at(start, end)
        row_elem = fromstring(_ROW_WRAP_PREFIX + chunk + _ROW_WRAP_SUFFIX)[0]
        parser = self._make_parser(None)
        parser.shared_formulae = self._shared_formulae
        # Ensure dependents without r= still get the right row counter if needed.
        parser.row_counter = row - 1
        return parser.parse_row(row_elem)

    def _get_cell(self, row, column):
        self._ensure_open()
        _idx, cells = self._parse_indexed_row(row)
        for cell_row, cell_col, cell_val, cell_dt, cell_sid, cell_cached in cells:
            if cell_col == column:
                return ReadOnlyCell(
                    self,
                    cell_row,
                    cell_col,
                    cell_val,
                    cell_dt,
                    cell_sid,
                    cached_value=cell_cached,
                )
        return EMPTY_CELL

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._row_index = {}
        self._shared_formulae = {}
        self._part.close()

    def _ensure_open(self) -> None:
        if self._closed:
            raise ValueError("I/O operation on closed IndexedWorksheet")
