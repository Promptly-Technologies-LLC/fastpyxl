# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors.excel import Relation


class Drawing(Serialisable):

    tagname = "drawing"

    id = Relation()

    def __init__(self, id=None):
        self.id = id
