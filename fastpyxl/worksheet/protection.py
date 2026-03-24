# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.utils.protection import hash_password


class _Protected:
    _password = None

    def set_password(self, value='', already_hashed=False):
        """Set a password on this sheet."""
        if not already_hashed:
            value = hash_password(value)
        self._password = value

    @property
    def password(self):
        """Return the password value, regardless of hash."""
        return self._password

    @password.setter
    def password(self, value):
        """Set a password directly, forcing a hash step."""
        self.set_password(value)


class SheetProtection(Serialisable, _Protected):
    """
    Information about protection of various aspects of a sheet. True values
    mean that protection for the object or action is active This is the
    **default** when protection is active, ie. users cannot do something
    """

    tagname = "sheetProtection"

    sheet: bool = Field.attribute(expected_type=bool, default=False)
    enabled = AliasField('sheet', default=None)
    objects: bool = Field.attribute(expected_type=bool, default=False)
    scenarios: bool = Field.attribute(expected_type=bool, default=False)
    formatCells: bool = Field.attribute(expected_type=bool, default=True)
    formatColumns: bool = Field.attribute(expected_type=bool, default=True)
    formatRows: bool = Field.attribute(expected_type=bool, default=True)
    insertColumns: bool = Field.attribute(expected_type=bool, default=True)
    insertRows: bool = Field.attribute(expected_type=bool, default=True)
    insertHyperlinks: bool = Field.attribute(expected_type=bool, default=True)
    deleteColumns: bool = Field.attribute(expected_type=bool, default=True)
    deleteRows: bool = Field.attribute(expected_type=bool, default=True)
    selectLockedCells: bool = Field.attribute(expected_type=bool, default=False)
    selectUnlockedCells: bool = Field.attribute(expected_type=bool, default=False)
    sort: bool = Field.attribute(expected_type=bool, default=True)
    autoFilter: bool = Field.attribute(expected_type=bool, default=True)
    pivotTables: bool = Field.attribute(expected_type=bool, default=True)
    saltValue: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    spinCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    algorithmName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    hashValue: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self, sheet=False, objects=False, scenarios=False,
                 formatCells=True, formatRows=True, formatColumns=True,
                 insertColumns=True, insertRows=True, insertHyperlinks=True,
                 deleteColumns=True, deleteRows=True, selectLockedCells=False,
                 selectUnlockedCells=False, sort=True, autoFilter=True, pivotTables=True,
                 password=None, algorithmName=None, saltValue=None, spinCount=None, hashValue=None):
        self.sheet = sheet
        self.objects = objects
        self.scenarios = scenarios
        self.formatCells = formatCells
        self.formatColumns = formatColumns
        self.formatRows = formatRows
        self.insertColumns = insertColumns
        self.insertRows = insertRows
        self.insertHyperlinks = insertHyperlinks
        self.deleteColumns = deleteColumns
        self.deleteRows = deleteRows
        self.selectLockedCells = selectLockedCells
        self.selectUnlockedCells = selectUnlockedCells
        self.sort = sort
        self.autoFilter = autoFilter
        self.pivotTables = pivotTables
        if password is not None:
            self.password = password
        self.algorithmName = algorithmName
        self.saltValue = saltValue
        self.spinCount = spinCount
        self.hashValue = hashValue

    def __iter__(self):
        ordered = (
            'selectLockedCells', 'selectUnlockedCells', 'algorithmName',
            'sheet', 'objects', 'insertRows', 'insertHyperlinks', 'autoFilter',
            'scenarios', 'formatColumns', 'deleteColumns', 'insertColumns',
            'pivotTables', 'deleteRows', 'formatCells', 'saltValue', 'formatRows',
            'sort', 'spinCount', 'password', 'hashValue',
        )
        for key in ordered:
            if key == 'password':
                value = self.password
            else:
                value = getattr(self, key, None)
            if value is None:
                continue
            yield key, safe_string(value)

    def set_password(self, value='', already_hashed=False):
        super().set_password(value, already_hashed)
        self.enable()

    def enable(self):
        self.sheet = True

    def disable(self):
        self.sheet = False

    def __bool__(self):
        return self.sheet
