# Copyright (c) 2010-2024 fastpyxl


# package imports
from fastpyxl.reader.strings import read_string_table
from fastpyxl.reader.strings import read_rich_text
from fastpyxl.cell.rich_text import TextBlock, CellRichText
from fastpyxl.cell.text import InlineFont
from fastpyxl.styles.colors import Color


def test_read_string_table(datadir):
    datadir.chdir()
    src = 'sharedStrings.xml'
    with open(src, "rb") as content:
        assert read_string_table(content) == [
                u'This is cell A1 in Sheet 1', u'This is cell G5']


def test_empty_string(datadir):
    datadir.chdir()
    src = 'sharedStrings-emptystring.xml'
    with open(src, "rb") as content:
        assert read_string_table(content) == [u'Testing empty cell', u'']


def test_formatted_string_table(datadir):
    datadir.chdir()
    src = 'shared-strings-rich.xml'
    expected = [
        'Welcome',
        CellRichText([
            'to the best ',
            TextBlock(
                font=InlineFont(
                    rFont='Calibri',
                    sz=11.0,
                    family=2.0,
                    scheme='minor',
                    color=Color(theme=1),
                    b=True,
                ),
                text='shop in ',
            ),
            TextBlock(
                font=InlineFont(
                    rFont='Calibri',
                    sz=11.0,
                    family=2.0,
                    scheme='minor',
                    color=Color(theme=1),
                    b=True,
                    u='single',
                ),
                text='town',
            ),
        ]),
        "     let's play ",
    ]
    with open(src, "rb") as content:
        assert read_rich_text(content) == expected
