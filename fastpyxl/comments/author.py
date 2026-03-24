# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field


class AuthorList(Serialisable):

    tagname = "authors"

    author: list[str] = Field.sequence(expected_type=str, default=list)
    authors = AliasField("author")

    def __init__(
        self,
        author=(),
    ):
        self.author = author
