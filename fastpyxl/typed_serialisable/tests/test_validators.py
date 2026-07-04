import pytest

from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.validators import hex_binary_converter


def test_hex_binary_converter_accepts_valid_hex():
    assert hex_binary_converter("a1B2c3") == "a1B2c3"


def test_hex_binary_converter_returns_none_for_none():
    assert hex_binary_converter(None) is None


@pytest.mark.parametrize("value", ["", "deadbeef!", "12 34", "0xFF"])
def test_hex_binary_converter_rejects_invalid(value):
    with pytest.raises(FieldValidationError, match="hex binary rejected"):
        hex_binary_converter(value)
