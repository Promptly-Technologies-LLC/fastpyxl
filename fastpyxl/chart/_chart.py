# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from collections import OrderedDict
from operator import attrgetter

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from ._3d import _3DBase
from .data_source import AxDataSource, NumRef
from .layout import Layout
from .legend import Legend
from .reference import Reference
from .series_factory import SeriesFactory
from .series import attribute_mapping
from .shapes import GraphicalProperties
from .title import title_from_value


def PlotArea():
    from .plotarea import PlotArea

    return PlotArea()


class ChartBase(Serialisable):
    """
    Base class for all charts
    """

    _plot_xml_tag: str | None = None

    display_blanks: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        default="gap",
        xml_name="display_blanks",
    )
    visible_cells_only: bool | None = Field.attribute(
        expected_type=bool,
        allow_none=True,
        default=True,
        xml_name="visible_cells_only",
    )

    axId: list[int] | None = Field.sequence(
        expected_type=int,
        allow_none=True,
        primitive_attribute="val",
        default=list,
    )

    _series_type = ""
    series = AliasField("ser")

    anchor = "E15"
    width = 15
    height = 7.5
    _id = 1
    _path = "/xl/charts/chart{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

    def __init__(self, axId=(), **kw):
        object.__init__(self)
        self._charts = [self]
        self._title = None
        self.layout = None
        self.roundedCorners = None
        self.legend = Legend()
        self.graphical_properties = None
        self.style = None
        self.plot_area = PlotArea()
        self.axId = list(axId)
        self.pivotSource = None
        self.pivotFormats = ()
        self.idx_base = 0
        self.display_blanks = "gap"
        self.visible_cells_only = True
        for key, value in kw.items():
            setattr(self, key, value)
        if not hasattr(self, "ser"):
            self.ser = []
        if self._axes and not self.axId:
            self.axId = list(self._axes.keys())

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        self._title = title_from_value(value)

    def __iter__(self):
        for attr in self.__attrs__:
            if type(self) is not ChartBase and attr in (
                "display_blanks",
                "visible_cells_only",
            ):
                continue
            field = self.__fields__[attr]
            value = getattr(self, attr)
            if value is None:
                continue
            xml_attr = field.tag
            if xml_attr.startswith("_"):
                xml_attr = xml_attr[1:]
            if field.hyphenated:
                xml_attr = xml_attr.replace("_", "-")
            if attr == "visible_cells_only":
                yield xml_attr, "1" if value else "0"
            else:
                yield xml_attr, safe_string(value)

    def __hash__(self):
        return id(self)

    def __iadd__(self, other):
        if not isinstance(other, ChartBase):
            raise TypeError("Only other charts can be added")
        self._charts.append(other)
        return self

    def to_tree(self, tagname=None, idx=None, namespace=None):
        del idx
        self.axId = [i for i in self._axes]
        if self.ser is not None:
            for s in self.ser:
                s.__elements__ = attribute_mapping[self._series_type]
        resolved = tagname
        if resolved is None:
            cls = type(self)
            own_tn = cls.__dict__.get("tagname")
            if isinstance(own_tn, str):
                resolved = own_tn
            elif cls is ChartBase:
                pt = getattr(ChartBase, "_plot_xml_tag", None)
                if pt:
                    resolved = pt
                else:
                    raise NotImplementedError
            else:
                resolved = self.tagname
        return super(ChartBase, self).to_tree(resolved, None, namespace)

    def _reindex(self):
        ds = sorted(self.series, key=attrgetter("order"))
        for idx, s in enumerate(ds):
            s.order = idx
        self.series = ds

    def _write(self):
        from .chartspace import ChartSpace, ChartContainer

        self.plot_area.layout = self.layout

        idx_base = self.idx_base
        for chart in self._charts:
            if chart not in self.plot_area._charts:
                chart.idx_base = idx_base
                idx_base += len(chart.series)
        self.plot_area._charts = self._charts

        chart = self._charts[-1]
        container = ChartContainer(
            plotArea=self.plot_area, legend=self.legend, title=self.title
        )
        if isinstance(chart, _3DBase):
            container.view3D = chart.view3D
            container.floor = chart.floor
            container.sideWall = chart.sideWall
            container.backWall = chart.backWall
        container.plotVisOnly = self.visible_cells_only
        container.dispBlanksAs = self.display_blanks
        container.pivotFmts = self.pivotFormats
        cs = ChartSpace(chart=container)
        cs.style = self.style
        cs.roundedCorners = self.roundedCorners
        cs.pivotSource = self.pivotSource
        cs.spPr = self.graphical_properties
        return cs.to_tree()

    @property
    def _axes(self):
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        return OrderedDict([(axis.axId, axis) for axis in (x, y, z) if axis])

    def set_categories(self, labels):
        if not isinstance(labels, Reference):
            labels = Reference(range_string=labels)
        for s in self.ser:
            s.cat = AxDataSource(numRef=NumRef(f=labels))

    def add_data(self, data, from_rows=False, titles_from_data=False):
        if not isinstance(data, Reference):
            data = Reference(range_string=data)

        if from_rows:
            values = data.rows
        else:
            values = data.cols

        for ref in values:
            series = SeriesFactory(ref, title_from_data=titles_from_data)
            self.series.append(series)

    def append(self, value):
        series_list = self.series[:]
        series_list.append(value)
        self.series = series_list

    @property
    def path(self):
        return self._path.format(self._id)
