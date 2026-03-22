# Copyright (c) 2010-2024 fastpyxl


from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors import (
    Sequence,
    Alias
)


class AuthorList(Serialisable):

    tagname = "authors"

    author = Sequence(expected_type=str)
    authors = Alias("author")

    def __init__(self,
                 author=(),
                ):
        self.author = author
