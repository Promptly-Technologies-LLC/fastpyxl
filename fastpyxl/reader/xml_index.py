# Copyright (c) 2010-2024 fastpyxl

"""Byte-offset indexes into decompressed SpreadsheetML parts."""

from __future__ import annotations

import re
from typing import Iterator


# Match an unprefixed start tag name boundary: <row ...>, <row/>, <si>, etc.
_ROW_START = re.compile(rb"<row(?=[\s/>])")
_SI_START = re.compile(rb"<si(?=[\s/>])")
_ROW_ATTR_R = re.compile(rb'\br="(\d+)"')
_SHEET_DATA_START = re.compile(rb"<sheetData(?=[\s/>])")
_SHEET_DATA_END = b"</sheetData>"


def _find_element_end(data: bytes, start: int, tag: bytes) -> int:
    """Return exclusive end offset of the element that starts at ``start``."""
    # Scan for end of start tag, tracking whether it is self-closing.
    i = start + 1 + len(tag)
    n = len(data)
    in_quote = None
    while i < n:
        ch = data[i]
        if in_quote:
            if ch == in_quote:
                in_quote = None
            i += 1
            continue
        if ch in (34, 39):  # " or '
            in_quote = ch
            i += 1
            continue
        if ch == 60:  # < inside start tag is malformed; stop defensively
            break
        if ch == 62:  # >
            # Self-closing if previous non-space was /
            j = i - 1
            while j > start and data[j] in b" \t\r\n":
                j -= 1
            if data[j] == 47:  # /
                return i + 1
            break
        i += 1
    else:
        raise ValueError(f"unclosed start tag <{tag.decode()}> at offset {start}")

    close = b"</" + tag + b">"
    end = data.find(close, i + 1)
    if end < 0:
        raise ValueError(f"missing close tag </{tag.decode()}> for offset {start}")
    return end + len(close)


def iter_row_spans(data: bytes) -> Iterator[tuple[int, int, int]]:
    """Yield ``(row_number, start, end)`` for each ``<row>`` in sheetData."""
    m = _SHEET_DATA_START.search(data)
    if m is None:
        return
    # Move past the <sheetData ...> start tag.
    pos = data.find(b">", m.start())
    if pos < 0:
        return
    pos += 1
    end_limit = data.find(_SHEET_DATA_END, pos)
    if end_limit < 0:
        end_limit = len(data)

    region = data
    next_row = 1
    while True:
        m = _ROW_START.search(region, pos, end_limit)
        if m is None:
            break
        start = m.start()
        end = _find_element_end(region, start, b"row")
        if end > end_limit:
            break
        # Prefer explicit r=; otherwise assign sequentially.
        header_end = region.find(b">", start, end)
        header = region[start : header_end + 1] if header_end >= 0 else region[start:end]
        rm = _ROW_ATTR_R.search(header)
        if rm:
            row_num = int(rm.group(1))
        else:
            row_num = next_row
        next_row = row_num + 1
        yield row_num, start, end
        pos = end


def build_row_index(data: bytes) -> dict[int, tuple[int, int]]:
    """Map logical row number → ``(start, end)`` exclusive byte offsets."""
    return {row: (start, end) for row, start, end in iter_row_spans(data)}


def build_si_index(data: bytes) -> list[tuple[int, int]]:
    """Return ordered ``(start, end)`` spans for each ``<si>`` element."""
    spans: list[tuple[int, int]] = []
    pos = 0
    n = len(data)
    while True:
        m = _SI_START.search(data, pos, n)
        if m is None:
            break
        start = m.start()
        end = _find_element_end(data, start, b"si")
        spans.append((start, end))
        pos = end
    return spans
