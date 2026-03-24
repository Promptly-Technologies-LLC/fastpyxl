import pytest

from fastpyxl.tests.helper import compare_xml
from fastpyxl.xml.functions import fromstring, tostring

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field


class AliasDemo(Serialisable):
    tagname = "aliasDemo"
    source: str | None = Field.nested_text(expected_type=str, allow_none=True, default=None)
    alias: str | None = AliasField("source", default=None)


class SequenceDemo(Serialisable):
    tagname = "sequenceDemo"
    value: list[int | str] = Field.sequence(expected_type=int, default=list)


class SequenceChild(Serialisable):
    tagname = "member"
    idx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)


class NestedSequenceDemo(Serialisable):
    tagname = "nestedSequenceDemo"
    groupMembers: list[SequenceChild] = Field.nested_sequence(
        expected_type=SequenceChild,
        count=True,
        default=list,
    )


def _validate_axis_mode(value: str | None) -> None:
    if value is None:
        return
    allowed = {"major", "minor", "none"}
    if value not in allowed:
        raise ValueError(f"mode rejected value {value!r}")


class NoneSetDemo(Serialisable):
    tagname = "noneSetDemo"
    mode: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        validator=_validate_axis_mode, default=None,
    )


def _convert_number_value(value: object) -> float | str | None:
    if value is None:
        return None
    if value == "#N/A":
        return "#N/A"
    return float(str(value))


class NumberValueDemo(Serialisable):
    tagname = "numVal"
    v: float | str | None = Field.nested_text(
        expected_type=object,
        allow_none=True,
        converter=_convert_number_value, default=None,
    )


class PartA(Serialisable):
    tagname = "a"
    v: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)


class PartB(Serialisable):
    tagname = "b"
    label: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)


class MultiSequenceDemo(Serialisable):
    tagname = "multiDemo"
    entries: list[PartA | PartB] = Field.multi_sequence(
        parts={"a": PartA, "b": PartB},
        default=list,
    )


class FontModel(Serialisable):
    tagname = "font"
    bold: bool | None = Field.nested_bool(allow_none=True, default=None)


class StyledCell(Serialisable):
    tagname = "cell"
    font: FontModel | None = Field.element(expected_type=FontModel, allow_none=True, default=None)


class WorkbookStyleRegistry:
    def __init__(self):
        self.fonts: list[FontModel] = []

    def register_cell_style(self, cell: StyledCell) -> None:
        if cell.font is None:
            return
        self.fonts.append(cell.font)


def test_alias_translation_forwards_access_without_extra_xml():
    model = AliasDemo(source="Sheet1!A1")
    assert model.alias == "Sheet1!A1"

    model.alias = "Sheet1!B2"
    assert model.source == "Sheet1!B2"

    xml = tostring(model.to_tree())
    expected = """
    <aliasDemo>
      <source>Sheet1!B2</source>
    </aliasDemo>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_sequence_translation_roundtrip():
    model = SequenceDemo(value=[1, "2", 3])
    xml = tostring(model.to_tree())
    expected = """
    <sequenceDemo>
      <value>1</value>
      <value>2</value>
      <value>3</value>
    </sequenceDemo>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = SequenceDemo.from_tree(fromstring(xml))
    assert parsed.value == [1, 2, 3]


def test_nested_sequence_translation_with_count_roundtrip():
    model = NestedSequenceDemo(groupMembers=[SequenceChild(idx=1), SequenceChild(idx=2)])
    xml = tostring(model.to_tree())
    expected = """
    <nestedSequenceDemo>
      <groupMembers count="2">
        <member idx="1"></member>
        <member idx="2"></member>
      </groupMembers>
    </nestedSequenceDemo>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = NestedSequenceDemo.from_tree(fromstring(xml))
    assert [m.idx for m in parsed.groupMembers] == [1, 2]


def test_noneset_translation_validator_enforces_allowed_values():
    model = NoneSetDemo(mode="major")
    assert model.mode == "major"

    model.mode = None
    assert model.mode is None

    with pytest.raises(ValueError, match="mode rejected value"):
        model.mode = "invalid"


def test_number_value_translation_preserves_na_and_numbers():
    na_model = NumberValueDemo(v="#N/A")
    assert na_model.v == "#N/A"

    number_model = NumberValueDemo(v="42.5")
    assert number_model.v == 42.5

    xml = tostring(NumberValueDemo(v="#N/A").to_tree())
    expected = """
    <numVal>
      <v>#N/A</v>
    </numVal>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = NumberValueDemo.from_tree(fromstring("<numVal><v>3.5</v></numVal>"))
    assert parsed.v == 3.5


def test_multisequence_translation_dispatches_by_tag():
    model = MultiSequenceDemo(entries=[PartA(v=1), PartB(label="x"), PartA(v=2)])
    xml = tostring(model.to_tree())
    expected = """
    <multiDemo>
      <a v="1"></a>
      <b label="x"></b>
      <a v="2"></a>
    </multiDemo>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = MultiSequenceDemo.from_tree(fromstring(xml))
    assert isinstance(parsed.entries[0], PartA)
    assert isinstance(parsed.entries[1], PartB)
    assert isinstance(parsed.entries[2], PartA)
    assert parsed.entries[1].label == "x"


def test_style_side_effect_playbook_uses_explicit_registration():
    registry = WorkbookStyleRegistry()
    cell = StyledCell(font=FontModel(bold=True))

    # Assignment remains side-effect free until writer/save registration.
    assert registry.fonts == []
    cell.font = FontModel(bold=False)
    assert registry.fonts == []

    registry.register_cell_style(cell)
    assert len(registry.fonts) == 1
    assert registry.fonts[0].bold is False
