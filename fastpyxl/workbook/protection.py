# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.utils.protection import hash_password


class WorkbookProtection(Serialisable):

    _workbook_password, _revisions_password = None, None

    tagname = "workbookPr"

    workbook_password = AliasField("workbookPassword")
    workbookPasswordCharacterSet: str | None = Field.attribute(expected_type=str, allow_none=True)
    revision_password = AliasField("revisionsPassword")
    revisionsPasswordCharacterSet: str | None = Field.attribute(expected_type=str, allow_none=True)
    lockStructure: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    lock_structure = AliasField("lockStructure")
    lockWindows: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    lock_windows = AliasField("lockWindows")
    lockRevision: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    lock_revision = AliasField("lockRevision")
    revisionsAlgorithmName: str | None = Field.attribute(expected_type=str, allow_none=True)
    revisionsHashValue: str | None = Field.attribute(expected_type=str, allow_none=True)
    revisionsSaltValue: str | None = Field.attribute(expected_type=str, allow_none=True)
    revisionsSpinCount: int | None = Field.attribute(expected_type=int, allow_none=True)
    workbookAlgorithmName: str | None = Field.attribute(expected_type=str, allow_none=True)
    workbookHashValue: str | None = Field.attribute(expected_type=str, allow_none=True)
    workbookSaltValue: str | None = Field.attribute(expected_type=str, allow_none=True)
    workbookSpinCount: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 workbookPassword=None,
                 workbookPasswordCharacterSet=None,
                 revisionsPassword=None,
                 revisionsPasswordCharacterSet=None,
                 lockStructure=None,
                 lockWindows=None,
                 lockRevision=None,
                 revisionsAlgorithmName=None,
                 revisionsHashValue=None,
                 revisionsSaltValue=None,
                 revisionsSpinCount=None,
                 workbookAlgorithmName=None,
                 workbookHashValue=None,
                 workbookSaltValue=None,
                 workbookSpinCount=None,
                ):
        if workbookPassword is not None:
            self.workbookPassword = workbookPassword
        self.workbookPasswordCharacterSet = workbookPasswordCharacterSet
        if revisionsPassword is not None:
            self.revisionsPassword = revisionsPassword
        self.revisionsPasswordCharacterSet = revisionsPasswordCharacterSet
        self.lockStructure = lockStructure
        self.lockWindows = lockWindows
        self.lockRevision = lockRevision
        self.revisionsAlgorithmName = revisionsAlgorithmName
        self.revisionsHashValue = revisionsHashValue
        self.revisionsSaltValue = revisionsSaltValue
        self.revisionsSpinCount = revisionsSpinCount
        self.workbookAlgorithmName = workbookAlgorithmName
        self.workbookHashValue = workbookHashValue
        self.workbookSaltValue = workbookSaltValue
        self.workbookSpinCount = workbookSpinCount

    def set_workbook_password(self, value='', already_hashed=False):
        """Set a password on this workbook."""
        if not already_hashed:
            value = hash_password(value)
        self._workbook_password = value

    @property
    def workbookPassword(self):
        """Return the workbook password value, regardless of hash."""
        return self._workbook_password

    @workbookPassword.setter
    def workbookPassword(self, value):
        """Set a workbook password directly, forcing a hash step."""
        self.set_workbook_password(value)

    def set_revisions_password(self, value='', already_hashed=False):
        """Set a revision password on this workbook."""
        if not already_hashed:
            value = hash_password(value)
        self._revisions_password = value

    @property
    def revisionsPassword(self):
        """Return the revisions password value, regardless of hash."""
        return self._revisions_password

    @revisionsPassword.setter
    def revisionsPassword(self, value):
        """Set a revisions password directly, forcing a hash step."""
        self.set_revisions_password(value)

    @classmethod
    def from_tree(cls, node):
        """Don't hash passwords when deserialising from XML"""
        self = super().from_tree(node)
        if self.workbookPassword:
            self.set_workbook_password(node.get('workbookPassword'), already_hashed=True)
        if self.revisionsPassword:
            self.set_revisions_password(node.get('revisionsPassword'), already_hashed=True)
        return self

# Backwards compatibility
DocumentSecurity = WorkbookProtection


class FileSharing(Serialisable):

    tagname = "fileSharing"

    readOnlyRecommended: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    userName: str | None = Field.attribute(expected_type=str, allow_none=True)
    reservationPassword: str | None = Field.attribute(expected_type=str, allow_none=True)
    algorithmName: str | None = Field.attribute(expected_type=str, allow_none=True)
    hashValue: str | None = Field.attribute(expected_type=str, allow_none=True)
    saltValue: str | None = Field.attribute(expected_type=str, allow_none=True)
    spinCount: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 readOnlyRecommended=None,
                 userName=None,
                 reservationPassword=None,
                 algorithmName=None,
                 hashValue=None,
                 saltValue=None,
                 spinCount=None,
                ):
        self.readOnlyRecommended = readOnlyRecommended
        self.userName = userName
        self.reservationPassword = reservationPassword
        self.algorithmName = algorithmName
        self.hashValue = hashValue
        self.saltValue = saltValue
        self.spinCount = spinCount
