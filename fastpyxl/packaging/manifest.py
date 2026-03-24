# Copyright (c) 2010-2024 fastpyxl

"""
File manifest
"""
from mimetypes import MimeTypes
import os.path

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.functions import fromstring
from fastpyxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_THEME,
    ARC_STYLE,
    THEME_TYPE,
    STYLES_TYPE,
    CONTYPES_NS,
    ACTIVEX,
    CTRL,
    VBA,
)
from fastpyxl.xml.functions import tostring

# initialise mime-types
mimetypes = MimeTypes()
mimetypes.add_type('application/xml', ".xml")
mimetypes.add_type('application/vnd.openxmlformats-package.relationships+xml', ".rels")
mimetypes.add_type("application/vnd.ms-office.vbaProject", ".bin")
mimetypes.add_type("application/vnd.openxmlformats-officedocument.vmlDrawing", ".vml")
mimetypes.add_type("image/x-emf", ".emf")


class FileExtension(Serialisable):

    tagname = "Default"

    Extension: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    ContentType: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self, Extension, ContentType):
        self.Extension = Extension
        self.ContentType = ContentType


class Override(Serialisable):

    tagname = "Override"

    PartName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    ContentType: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self, PartName, ContentType):
        self.PartName = PartName
        self.ContentType = ContentType


ContentTypeOverride = Override


DEFAULT_TYPES = [
    FileExtension("rels", "application/vnd.openxmlformats-package.relationships+xml"),
    FileExtension("xml", "application/xml"),
]

DEFAULT_OVERRIDE = [
    Override("/" + ARC_STYLE, STYLES_TYPE), # Styles
    Override("/" + ARC_THEME, THEME_TYPE), # Theme
    Override("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    Override("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
]


class _AppendUniqueList(list):
    def append(self, value):
        if value in self:
            return
        super().append(value)


class Manifest(Serialisable):

    tagname = "Types"

    Default: list[FileExtension] = Field.sequence(
        expected_type=FileExtension, default=list, container_factory=_AppendUniqueList
    )
    Override: list[Override] = Field.sequence(
        expected_type=Override, default=list, container_factory=_AppendUniqueList
    )
    path = "[Content_Types].xml"

    xml_order = ("Default", "Override")

    def __init__(self,
                 Default=(),
                 Override=(),
                 ):
        if not Default:
            Default = DEFAULT_TYPES
        self.Default = Default
        if not Override:
            Override = DEFAULT_OVERRIDE
        self.Override = Override


    @property
    def filenames(self):
        return [part.PartName for part in self.Override]


    @property
    def extensions(self):
        """
        Map content types to file extensions
        Skip parts without extensions
        """
        exts = {os.path.splitext(str(part.PartName))[-1] for part in self.Override}
        return [(ext[1:], mimetypes.types_map[True][ext]) for ext in sorted(exts) if ext]


    def to_tree(self):
        """
        Custom serialisation method to allow setting a default namespace
        """
        defaults = [t.Extension for t in self.Default]
        for ext, mime in self.extensions:
            if ext not in defaults:
                mime = FileExtension(ext, mime)
                self._append_default_if_missing(mime)
        tree = super().to_tree()
        tree.set("xmlns", CONTYPES_NS)
        return tree


    def __contains__(self, content_type):
        """
        Check whether a particular content type is contained
        """
        for t in self.Override:
            if t.ContentType == content_type:
                return True


    def find(self, content_type):
        """
        Find specific content-type
        """
        try:
            return next(self.findall(content_type))
        except StopIteration:
            return


    def findall(self, content_type):
        """
        Find all elements of a specific content-type
        """
        for t in self.Override:
            if t.ContentType == content_type:
                yield t


    def append(self, obj):
        """
        Add content object to the package manifest
        # needs a contract...
        """
        ct = Override(PartName=obj.path, ContentType=obj.mime_type)
        self._append_override_if_missing(ct)


    def _write(self, archive, workbook):
        """
        Write manifest to the archive
        """
        self.append(workbook)
        self._write_vba(workbook)
        self._register_mimetypes(filenames=archive.namelist())
        archive.writestr(self.path, tostring(self.to_tree()))


    def _register_mimetypes(self, filenames):
        """
        Make sure that the mime type for all file extensions is registered
        """
        for fn in filenames:
            ext = os.path.splitext(fn)[-1]
            if not ext:
                continue
            mime = mimetypes.types_map[True][ext]
            fe = FileExtension(ext[1:], mime)
            self._append_default_if_missing(fe)


    def _write_vba(self, workbook):
        """
        Add content types from cached workbook when keeping VBA
        """
        if workbook.vba_archive:
            node = fromstring(workbook.vba_archive.read(ARC_CONTENT_TYPES))
            mf = Manifest.from_tree(node)
            filenames = self.filenames
            for override in mf.Override:
                if override.PartName not in (ACTIVEX, CTRL, VBA):
                    continue
                if override.PartName not in filenames:
                    self._append_override_if_missing(override)


    def _append_default_if_missing(self, item: FileExtension):
        for existing in self.Default:
            if existing.Extension == item.Extension and existing.ContentType == item.ContentType:
                return
        self.Default.append(item)


    def _append_override_if_missing(self, item: ContentTypeOverride):
        for existing in self.Override:
            if existing.PartName == item.PartName and existing.ContentType == item.ContentType:
                return
        self.Override.append(item)
