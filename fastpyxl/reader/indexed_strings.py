# Copyright (c) 2010-2024 fastpyxl

"""On-demand shared string table backed by a byte-offset index."""

from __future__ import annotations

from collections import OrderedDict
from typing import TYPE_CHECKING

from fastpyxl.cell.rich_text import CellRichText
from fastpyxl.cell.text import Text
from fastpyxl.xml.constants import SHEET_MAIN_NS
from fastpyxl.xml.functions import fromstring

from .xml_index import build_si_index

if TYPE_CHECKING:
    from .part_cache import DecompressedPart


_SI_WRAP_PREFIX = (
    f'<sst xmlns="{SHEET_MAIN_NS}">'.encode("ascii")
)
_SI_WRAP_SUFFIX = b"</sst>"


def decode_si_bytes(si_xml: bytes, *, rich_text: bool = False):
    """Decode a single ``<si>...</si>`` fragment."""
    node = fromstring(_SI_WRAP_PREFIX + si_xml + _SI_WRAP_SUFFIX)[0]
    if rich_text:
        text = CellRichText.from_tree(node)
        if len(text) == 0:
            return ""
        if len(text) == 1 and isinstance(text[0], str):
            return text[0]
        return text
    text = Text.from_tree(node).content
    return text.replace("x005F_", "")


class IndexedSharedStrings:
    """Sequence-like SST that decodes ``<si>`` entries on demand."""

    __slots__ = ("_part", "_spans", "_rich_text", "_cache", "_cache_size", "_closed")

    def __init__(
        self,
        part: DecompressedPart,
        spans: list[tuple[int, int]] | None = None,
        *,
        rich_text: bool = False,
        cache_size: int = 256,
    ):
        self._part = part
        if spans is None:
            # Index from an in-memory snapshot when the part may be file-backed.
            with part.open() as src:
                data = src.read()
            spans = build_si_index(data)
        self._spans = spans
        self._rich_text = rich_text
        self._cache: OrderedDict[int, object] = OrderedDict()
        self._cache_size = cache_size
        self._closed = False

    @classmethod
    def from_part(cls, part: DecompressedPart, *, rich_text: bool = False, cache_size: int = 256):
        with part.open() as src:
            data = src.read()
        return cls(part, build_si_index(data), rich_text=rich_text, cache_size=cache_size)

    def __len__(self) -> int:
        return len(self._spans)

    def __iter__(self):
        for i in range(len(self)):
            yield self[i]

    def __getitem__(self, idx):
        self._ensure_open()
        if isinstance(idx, slice):
            return [self[i] for i in range(*idx.indices(len(self)))]
        if not isinstance(idx, int):
            raise TypeError("indices must be integers")
        if idx < 0:
            idx += len(self._spans)
        if idx < 0 or idx >= len(self._spans):
            raise IndexError("shared string index out of range")
        cached = self._cache.get(idx)
        if cached is not None:
            self._cache.move_to_end(idx)
            return cached
        start, end = self._spans[idx]
        value = decode_si_bytes(self._part.read_at(start, end), rich_text=self._rich_text)
        self._cache[idx] = value
        if len(self._cache) > self._cache_size:
            self._cache.popitem(last=False)
        return value

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._cache.clear()
        self._spans = []
        self._part.close()

    def _ensure_open(self) -> None:
        if self._closed:
            raise ValueError("I/O operation on closed IndexedSharedStrings")
