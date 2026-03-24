from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

from fastpyxl.xml.constants import REL_NS


class Hyperlink(Serialisable):

    tagname = "hyperlink"

    ref: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    location: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    tooltip: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    display: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    id: str | None = Field.attribute(
        expected_type=str, allow_none=True, namespace=REL_NS, default=None
    )

    def __init__(
        self,
        ref=None,
        location=None,
        tooltip=None,
        display=None,
        id=None,
        target=None,
    ):
        self.ref = ref
        self.location = location
        self.tooltip = tooltip
        self.display = display
        self.id = id
        self.target = target


class HyperlinkList(Serialisable):

    tagname = "hyperlinks"

    hyperlink: list[Hyperlink] = Field.sequence(expected_type=Hyperlink, default=list)

    def __init__(self, hyperlink=()):
        self.hyperlink = list(hyperlink)
