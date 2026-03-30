# Copyright (c) 2010-2024 fastpyxl

import re
import warnings

from fastpyxl.worksheet.header_footer import HeaderFooter

"""
Base class for worksheets, chartsheets, etc. that can be added to workbooks
"""

INVALID_TITLE_REGEX = re.compile(r'[\\*?:/\[\]]')


def avoid_duplicate_name(names, value):
    """
    Naive check to see whether name already exists.
    If name does exist suggest a name using an incrementer
    Duplicates are case insensitive
    """
    # Check for an absolute match in which case we need to find an alternative
    match = [n for n in names if n.lower() == value.lower()]
    if match:
        names = u",".join(names)
        sheet_title_regex = re.compile(f'(?P<title>{re.escape(value)})(?P<count>\\d*),?', re.I)
        matches = sheet_title_regex.findall(names)
        if matches:
            # use name, but append with the next highest integer
            counts = [int(idx) for (t, idx) in matches if idx.isdigit()]
            highest = 0
            if counts:
                highest = max(counts)
            value = u"{0}{1}".format(value, highest + 1)
    return value


class _WorkbookChild:

    __title = ""
    _id = None
    _path = "{0}"
    _parent = None
    _default_title = "Sheet"

    def __init__(self, parent=None, title=None):
        self._parent = parent
        self.title = title or self._default_title
        self.HeaderFooter = HeaderFooter()


    def __repr__(self):
        return '<{0} "{1}">'.format(self.__class__.__name__, self.title)


    @property
    def parent(self):
        return self._parent


    @property
    def encoding(self):
        parent = self._parent
        assert parent is not None
        return parent.encoding


    @property
    def title(self):
        return self.__title


    @title.setter
    def title(self, value):
        """
        Set a sheet title, ensuring it is valid.
        Limited to 31 characters, no special characters.
        Duplicate titles will be incremented numerically
        """
        if not self._parent:
            return

        if not value:
            raise ValueError("Title must have at least one character")

        if hasattr(value, "decode"):
            if not isinstance(value, str):
                try:
                    value = value.decode("ascii")
                except UnicodeDecodeError:
                    raise ValueError("Worksheet titles must be str")

        m = INVALID_TITLE_REGEX.search(value)
        if m:
            msg = "Invalid character {0} found in sheet title".format(m.group(0))
            raise ValueError(msg)

        parent = self._parent
        old_title = self.__title

        if old_title is not None and old_title != value:
            idx = getattr(parent, "_sheet_titles_lower", None)
            if idx is not None:
                vl = value.lower()
                conflict = vl in idx
                if self in parent._sheets and old_title:
                    if conflict and vl == old_title.lower():
                        conflict = False
                if conflict:
                    value = avoid_duplicate_name(parent.sheetnames, value)
            else:
                value = avoid_duplicate_name(parent.sheetnames, value)

        if len(value) > 31:
            warnings.warn("Title is more than 31 characters. Some applications may not be able to read the file")

        self.__title = value
        idx = getattr(parent, "_sheet_titles_lower", None)
        if idx is not None and self in parent._sheets:
            if old_title:
                idx.discard(old_title.lower())
            idx.add(value.lower())
            parent._sheetnames_cache = None
            parent._sheet_title_map = None


    @property
    def oddHeader(self):
        return self.HeaderFooter.oddHeader


    @oddHeader.setter
    def oddHeader(self, value):
        self.HeaderFooter.oddHeader = value


    @property
    def oddFooter(self):
        return self.HeaderFooter.oddFooter


    @oddFooter.setter
    def oddFooter(self, value):
        self.HeaderFooter.oddFooter = value


    @property
    def evenHeader(self):
        return self.HeaderFooter.evenHeader


    @evenHeader.setter
    def evenHeader(self, value):
        self.HeaderFooter.evenHeader = value


    @property
    def evenFooter(self):
        return self.HeaderFooter.evenFooter


    @evenFooter.setter
    def evenFooter(self, value):
        self.HeaderFooter.evenFooter = value


    @property
    def firstHeader(self):
        return self.HeaderFooter.firstHeader


    @firstHeader.setter
    def firstHeader(self, value):
        self.HeaderFooter.firstHeader = value


    @property
    def firstFooter(self):
        return self.HeaderFooter.firstFooter


    @firstFooter.setter
    def firstFooter(self, value):
        self.HeaderFooter.firstFooter = value


    @property
    def path(self):
        return self._path.format(self._id)
