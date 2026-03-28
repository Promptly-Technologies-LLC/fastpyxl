# Copyright (c) 2010-2024 fastpyxl

from copy import copy

from .numbers import (
    BUILTIN_FORMATS,
    BUILTIN_FORMATS_MAX_SIZE,
    BUILTIN_FORMATS_REVERSE,
)
from .proxy import StyleProxy
from .cell_style import StyleArray
from .named_styles import NamedStyle
from .builtins import styles


class StyleDescriptor:

    def __init__(self, collection, key):
        self.collection = collection
        self.key = key

    def __set__(self, instance, value):
        pending_styles = getattr(instance, "_pending_styles", None)
        if pending_styles is not None:
            instance._ensure_style_array()
            pending_styles[self.key] = (self.collection, value)
            return
        coll = getattr(instance.parent.parent, self.collection)
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, coll.add(value))


    def __get__(self, instance, cls):
        if instance is None:
            return self
        pending_styles = getattr(instance, "_pending_styles", None)
        if pending_styles is not None:
            pending = pending_styles.get(self.key)
            if pending is not None:
                _, value = pending
                return StyleProxy(value)
        coll = getattr(instance.parent.parent, self.collection)
        if hasattr(instance, "_ensure_style_array"):
            instance._ensure_style_array()
        elif not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        return StyleProxy(coll[idx])


class NumberFormatDescriptor:

    key = "numFmtId"
    collection = '_number_formats'

    def __set__(self, instance, value):
        pending_styles = getattr(instance, "_pending_styles", None)
        if pending_styles is not None:
            instance._ensure_style_array()
            pending_styles[self.key] = (self.collection, value)
            return
        coll = getattr(instance.parent.parent, self.collection)
        if value in BUILTIN_FORMATS_REVERSE:
            idx = BUILTIN_FORMATS_REVERSE[value]
        else:
            idx = coll.add(value) + BUILTIN_FORMATS_MAX_SIZE
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, idx)


    def __get__(self, instance, cls):
        if instance is None:
            return self
        pending_styles = getattr(instance, "_pending_styles", None)
        if pending_styles is not None:
            pending = pending_styles.get(self.key)
            if pending is not None:
                _, value = pending
                return value
        if hasattr(instance, "_ensure_style_array"):
            instance._ensure_style_array()
        elif not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        if idx < BUILTIN_FORMATS_MAX_SIZE:
            return BUILTIN_FORMATS.get(idx, "General")
        coll = getattr(instance.parent.parent, self.collection)
        return coll[idx - BUILTIN_FORMATS_MAX_SIZE]


class NamedStyleDescriptor:

    key = "xfId"
    collection = "_named_styles"


    def __set__(self, instance, value):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        wb = instance.parent.parent
        coll = getattr(wb, self.collection)
        pending_styles = getattr(instance, "_pending_styles", None)
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
        pending = getattr(instance, "_pending_named_style", None)
        if pending is not None:
            if isinstance(pending, NamedStyle):
                return pending.name
            return pending
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        coll = getattr(instance.parent.parent, self.collection)
        return coll.names[idx]


class StyleArrayDescriptor:

    def __init__(self, key):
        self.key = key

    def __set__(self, instance, value):
        if instance._style is None:
            instance._style = StyleArray()
        setattr(instance._style, self.key, value)


    def __get__(self, instance, cls):
        if instance._style is None:
            return False
        return bool(getattr(instance._style, self.key))


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

    __slots__ = ('parent', '_style', '_pending_styles', '_pending_named_style')

    def __init__(self, sheet, style_array=None):
        self.parent = sheet
        if style_array is not None:
            style_array = StyleArray(style_array)
        self._style = style_array
        self._pending_styles = {}
        self._pending_named_style = None

    def _ensure_style_array(self):
        if self._style is None:
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
        wb = self.parent.parent
        for key, (collection, value) in self._pending_styles.items():
            if key == "numFmtId":
                if value in BUILTIN_FORMATS_REVERSE:
                    idx = BUILTIN_FORMATS_REVERSE[value]
                else:
                    coll = getattr(wb, collection)
                    idx = coll.add(value) + BUILTIN_FORMATS_MAX_SIZE
            else:
                coll = getattr(wb, collection)
                idx = coll.add(value)
            setattr(self._style, key, idx)
        self._pending_styles.clear()


    @property
    def style_id(self):
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
            return False
        return any(self._style)


    def _style_id_for_save(self):
        """Return the style ID string if the cell is styled, else ``None``.

        Combines the checks of :attr:`has_style` with the materialisation
        performed by :attr:`style_id` so that the save path only traverses
        the style resolution chain once per cell.

        When ``_pending_styles`` is ``None`` (set by
        ``materialize_pending_style_components`` during save) all pending
        state has already been resolved, so we skip straight to the style
        array check.
        """
        # Fast path after pre-materialization: no pending state to resolve.
        if self._pending_styles is None:
            if self._style is None or not any(self._style):
                return None
            return str(self.parent.parent._cell_styles.add(self._style))
        # Slow path: pending styles still need resolution.
        if self._pending_named_style is not None:
            if self._style is None:
                self._style = StyleArray()
            self._apply_pending_named_style()
            self._apply_pending_styles()
            return str(self.parent.parent._cell_styles.add(self._style))
        if self._pending_styles:
            if self._style is None:
                self._style = StyleArray()
            self._apply_pending_styles()
            return str(self.parent.parent._cell_styles.add(self._style))
        if self._style is None:
            return None
        if not any(self._style):
            return None
        return str(self.parent.parent._cell_styles.add(self._style))

