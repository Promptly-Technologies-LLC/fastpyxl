# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors.excel import (
    Relation,
)

class ExternalReference(Serialisable):

    tagname = "externalReference"

    id = Relation()

    def __init__(self, id):
        self.id = id
