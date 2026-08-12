# Copyright (c) 2010-2024 fastpyxl

"""Seekable decompressed ZIP members for indexed reads."""

from __future__ import annotations

import os
import tempfile
from io import BytesIO


# Spill decompressed members larger than this to a temp file.
DEFAULT_SPOOL_MAX = 16 * 1024 * 1024


class DecompressedPart:
    """Immutable decompressed package part with random-access slices."""

    __slots__ = ("_buf", "_path", "_fd", "_size", "_closed")

    def __init__(self, data: bytes, *, spool_max: int = DEFAULT_SPOOL_MAX):
        self._closed = False
        self._size = len(data)
        if self._size <= spool_max:
            self._buf: bytes | None = data
            self._path: str | None = None
            self._fd = None
        else:
            self._buf = None
            fd, path = tempfile.mkstemp(prefix="fastpyxl-part-")
            self._fd = fd
            self._path = path
            os.write(fd, data)
            os.fsync(fd)

    @classmethod
    def from_archive(cls, archive, name: str, *, spool_max: int = DEFAULT_SPOOL_MAX):
        with archive.open(name) as src:
            data = src.read()
        return cls(data, spool_max=spool_max)

    @property
    def size(self) -> int:
        return self._size

    def read_at(self, start: int, end: int) -> bytes:
        self._ensure_open()
        if start < 0 or end < start or end > self._size:
            raise ValueError(f"invalid slice [{start}:{end}] for size {self._size}")
        if self._buf is not None:
            return self._buf[start:end]
        assert self._fd is not None
        os.lseek(self._fd, start, os.SEEK_SET)
        return os.read(self._fd, end - start)

    def open(self):
        """Return a fresh binary stream positioned at the start."""
        self._ensure_open()
        if self._buf is not None:
            return BytesIO(self._buf)
        assert self._path is not None
        return open(self._path, "rb")

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        self._buf = None
        if self._fd is not None:
            os.close(self._fd)
            self._fd = None
        if self._path is not None:
            try:
                os.unlink(self._path)
            except OSError:
                pass
            self._path = None

    def _ensure_open(self) -> None:
        if self._closed:
            raise ValueError("I/O operation on closed DecompressedPart")
