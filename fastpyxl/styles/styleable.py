# Copyright (c) 2010-2024 fastpyxl

from copy import copy
from operator import attrgetter

from .numbers import (
    BUILTIN_FORMATS,
    BUILTIN_FORMATS_MAX_SIZE,
    BUILTIN_FORMATS_REVERSE,
)
from .proxy import StyleProxy
from .cell_style import StyleArray
from .named_styles import NamedStyle
from .builtins import styles


# Map StyleArray attribute names to their underlying array indices.
_STYLE_KEY_INDEX = {
    "fontId": 0, "fillId": 1, "borderId": 2, "numFmtId": 3,
    "protectionId": 4, "alignmentId": 5, "pivotButton": 6,
    "quotePrefix": 7, "xfId": 8,
}

# Precomputed C-level attribute getters for workbook style collections.
_WB_COLLECTION_GETTER = {
    '_fonts': attrgetter('_fonts'),
    '_fills': attrgetter('_fills'),
    '_borders': attrgetter('_borders'),
    '_number_formats': attrgetter('_number_formats'),
    '_protections': attrgetter('_protections'),
    '_alignments': attrgetter('_alignments'),
    '_named_styles': attrgetter('_named_styles'),
}


class StyleDescriptor:

    def __init__(self, collection, key):
        self.collection = collection
        self.key = key
        self._key_idx = _STYLE_KEY_INDEX[key]
        self._get_collection = _WB_COLLECTION_GETTER[collection]

    def __set__(self, instance, value):
        pending_styles = instance._pending_styles
        if pending_styles is not None:
            instance._ensure_style_array()
            pending_styles[self.key] = (self.collection, value)
            return
        coll = self._get_collection(instance.parent.parent)
        if not instance._style:
            instance._style = StyleArray()
        instance._style[self._key_idx] = coll.add(value)


    def __get__(self, instance, cls):
        if instance is None:
            return self
        pending_styles = instance._pending_styles
        if pending_styles is not None:
            pending = pending_styles.get(self.key)
            if pending is not None:
                _, value = pending
                return StyleProxy(value)
        coll = self._get_collection(instance.parent.parent)
        if hasattr(instance, "_ensure_style_array"):
            instance._ensure_style_array()
        elif not instance._style:
            instance._style = StyleArray()
        idx = instance._style[self._key_idx]
        return StyleProxy(coll[idx])


class NumberFormatDescriptor:

    key = "numFmtId"
    collection = '_number_formats'
    _key_idx = _STYLE_KEY_INDEX["numFmtId"]
    _get_collection = _WB_COLLECTION_GETTER['_number_formats']

    def __set__(self, instance, value):
        pending_styles = instance._pending_styles
        if pending_styles is not None:
            instance._ensure_style_array()
            pending_styles[self.key] = (self.collection, value)
            return
        coll = self._get_collection(instance.parent.parent)
        if value in BUILTIN_FORMATS_REVERSE:
            idx = BUILTIN_FORMATS_REVERSE[value]
        else:
            idx = coll.add(value) + BUILTIN_FORMATS_MAX_SIZE
        if not instance._style:
            instance._style = StyleArray()
        instance._style[self._key_idx] = idx


    def __get__(self, instance, cls):
        if instance is None:
            return self
        pending_styles = instance._pending_styles
        if pending_styles is not None:
            pending = pending_styles.get(self.key)
            if pending is not None:
                _, value = pending
                return value
        if hasattr(instance, "_ensure_style_array"):
            instance._ensure_style_array()
        elif not instance._style:
            instance._style = StyleArray()
        idx = instance._style[self._key_idx]
        if idx < BUILTIN_FORMATS_MAX_SIZE:
            return BUILTIN_FORMATS.get(idx, "General")
        coll = self._get_collection(instance.parent.parent)
        return coll[idx - BUILTIN_FORMATS_MAX_SIZE]


class NamedStyleDescriptor:

    key = "xfId"
    collection = "_named_styles"
    _key_idx = _STYLE_KEY_INDEX["xfId"]
    _get_collection = _WB_COLLECTION_GETTER['_named_styles']


    def __set__(self, instance, value):
        instance._ensure_style_array()
        wb = instance.parent.parent
        coll = self._get_collection(wb)
        pending_styles = instance._pending_styles
        if pending_styles is not None:
            if isinstance(value, NamedStyle):
                if value in coll:
                    style = value
                    instance._style = copy(style.as_tuple())
                    instance._pending_named_style = None
                    return
                instance._pending_named_style = value
                return
            if value not in coll.names:
                if value in styles:
                    instance._pending_named_style = value
                    return
                raise ValueError("{0} is not a known style".format(value))
            instance._pending_named_style = value
            return
        if isinstance(value, NamedStyle):
            style = value
            if style not in coll:
                wb.add_named_style(style)
        elif value not in coll.names:
            if value in styles: # is it builtin?
                style = styles[value]
                if style not in coll:
                    wb.add_named_style(style)
            else:
                raise ValueError("{0} is not a known style".format(value))
        else:
            style = coll[value]
        instance._style = copy(style.as_tuple())
        instance._pending_named_style = None


    def __get__(self, instance, cls):
        if instance is None:
            return self
        pending = instance._pending_named_style
        if pending is not None:
            if isinstance(pending, NamedStyle):
                return pending.name
            return pending
        instance._ensure_style_array()
        idx = instance._style[self._key_idx]
        coll = self._get_collection(instance.parent.parent)
        return coll.names[idx]


class StyleArrayDescriptor:

    def __init__(self, key):
        self.key = key
        self._key_idx = _STYLE_KEY_INDEX[key]

    def __set__(self, instance, value):
        instance._ensure_style_array()
        instance._style[self._key_idx] = value


    def __get__(self, instance, cls):
        if instance._style is None:
            if not instance._style_id:
                return False
            instance._ensure_style_array()
        return bool(instance._style[self._key_idx])


class StyleableObject:
    """
    Base class for styleble objects implementing proxy and lookup functions
    """

    font = StyleDescriptor('_fonts', "fontId")
    fill = StyleDescriptor('_fills', "fillId")
    border = StyleDescriptor('_borders', "borderId")
    number_format = NumberFormatDescriptor()
    protection = StyleDescriptor('_protections', "protectionId")
    alignment = StyleDescriptor('_alignments', "alignmentId")
    style = NamedStyleDescriptor()
    quotePrefix = StyleArrayDescriptor('quotePrefix')
    pivotButton = StyleArrayDescriptor('pivotButton')

    __slots__ = ('parent', '_style', '_style_id', '_pending_styles', '_pending_named_style')

    def __init__(self, sheet, style_array=None):
        self.parent = sheet
        if style_array is not None:
            style_array = StyleArray(style_array)
        self._style = style_array
        self._style_id = 0
        self._pending_styles = {}
        self._pending_named_style = None

    def _ensure_style_array(self):
        if self._style is None:
            style_id = self._style_id
            if style_id:
                self._style = StyleArray(self.parent.parent._cell_styles[style_id])
            else:
                self._style = StyleArray()

    def _apply_pending_named_style(self):
        pending = self._pending_named_style
        if pending is None:
            return
        wb = self.parent.parent
        coll = wb._named_styles
        if isinstance(pending, NamedStyle):
            if pending.name in coll.names:
                style = coll[pending.name]
            else:
                wb.add_named_style(pending)
                style = pending
        elif pending in coll.names:
            style = coll[pending]
        elif pending in styles:
            st = styles[pending]
            if st.name in coll.names:
                style = coll[st.name]
            else:
                wb.add_named_style(st)
                style = st
        else:
            raise ValueError("{0} is not a known style".format(pending))
        self._style = copy(style.as_tuple())
        self._pending_named_style = None


    def _apply_pending_styles(self):
        if not self._pending_styles:
            return
        assert self._style is not None
        wb = self.parent.parent
        for key, (collection, value) in self._pending_styles.items():
            if key == "numFmtId":
                if value in BUILTIN_FORMATS_REVERSE:
                    idx = BUILTIN_FORMATS_REVERSE[value]
                else:
                    coll = _WB_COLLECTION_GETTER[collection](wb)
                    idx = coll.add(value) + BUILTIN_FORMATS_MAX_SIZE
            else:
                coll = _WB_COLLECTION_GETTER[collection](wb)
                idx = coll.add(value)
            self._style[_STYLE_KEY_INDEX[key]] = idx
        self._pending_styles.clear()


    @property
    def style_id(self):
        if self._style is None and self._style_id and not self._pending_styles and self._pending_named_style is None:
            return self._style_id
        self._ensure_style_array()
        self._apply_pending_named_style()
        self._apply_pending_styles()
        return self.parent.parent._cell_styles.add(self._style)


    @property
    def has_style(self):
        if self._pending_named_style is not None:
            return True
        if self._pending_styles:
            return True
        if self._style is None:
            return self._style_id != 0
        return any(self._style)
