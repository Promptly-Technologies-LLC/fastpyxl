# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.xml.constants import CHART_NS

from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors.excel import Relation


class ChartRelation(Serialisable):

    tagname = "chart"
    namespace = CHART_NS

    id = Relation()

    def __init__(self, id):
        self.id = id
