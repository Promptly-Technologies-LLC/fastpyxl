# Copyright (c) 2010-2024 fastpyxl

"""Seekable decompressed ZIP members for indexed reads."""

from __future__ import annotations

import mmap
import os
import tempfile
from contextlib import contextmanager
from io import BytesIO


# Spill decompressed members larger than this to a temp file.
DEFAULT_SPOOL_MAX = 16 * 1024 * 1024
_STREAM_CHUNK = 1024 * 1024


class DecompressedPart:
    """Immutable decompressed package part with random-access slices.

    Members larger than ``spool_max`` are streamed to a temporary file so the
    full payload is not held as a Python ``bytes`` object. Index construction
    should use :meth:`bytes_view` (mmap for spooled parts) rather than
    ``open().read()``, which would re-materialize the whole part in RSS.
    """

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
    def _from_spooled_file(cls, fd: int, path: str, size: int):
        part = object.__new__(cls)
        part._closed = False
        part._buf = None
        part._fd = fd
        part._path = path
        part._size = size
        return part

    @classmethod
    def from_archive(cls, archive, name: str, *, spool_max: int = DEFAULT_SPOOL_MAX):
        """Inflate a ZIP member, streaming large parts straight to a temp file."""
        info = archive.getinfo(name)
        # ZipInfo.file_size is the uncompressed size when known.
        uncompressed = getattr(info, "file_size", None)
        with archive.open(name) as src:
            if uncompressed is not None and uncompressed > spool_max:
                fd, path = tempfile.mkstemp(prefix="fastpyxl-part-")
                size = 0
                try:
                    while True:
                        chunk = src.read(_STREAM_CHUNK)
                        if not chunk:
                            break
                        os.write(fd, chunk)
                        size += len(chunk)
                    os.fsync(fd)
                except Exception:
                    os.close(fd)
                    try:
                        os.unlink(path)
                    except OSError:
                        pass
                    raise
                return cls._from_spooled_file(fd, path, size)
            data = src.read()
        # If the archive lied about size (or omitted it), still honor spool_max.
        return cls(data, spool_max=spool_max)

    @property
    def size(self) -> int:
        return self._size

    @property
    def spooled(self) -> bool:
        """True when the decompressed payload lives on disk, not in ``bytes``."""
        return self._buf is None and self._path is not None

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

    @contextmanager
    def bytes_view(self):
        """Yield a bytes-like view for scanning/indexing without a second copy.

        In-memory parts yield the underlying ``bytes``. Spooled parts yield an
        ``mmap`` window over the temp file.
        """
        self._ensure_open()
        if self._buf is not None:
            yield self._buf
            return
        assert self._path is not None
        with open(self._path, "rb") as fh:
            with mmap.mmap(fh.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                yield mm

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
