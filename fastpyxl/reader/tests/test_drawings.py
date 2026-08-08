# Copyright (c) 2010-2024 fastpyxl

import warnings
from io import BytesIO
from zipfile import ZipFile


def test_read_charts(datadir):
    datadir.chdir()

    archive = ZipFile("sample.xlsx")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    charts = find_images(archive, path)[0]
    assert len(charts) == 6


def test_read_drawing(datadir):
    datadir.chdir()

    archive = ZipFile("sample_with_images.xlsx")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    images = find_images(archive, path)[1]
    assert len(images) == 3


def test_unsupport_drawing(datadir):
    datadir.chdir()
    out = BytesIO()
    archive = ZipFile(out, mode="w")
    archive.write("unsupported_drawing.xml", "drawing1.xml")

    from ..drawings import find_images
    charts, images = find_images(archive, "drawing1.xml")
    assert charts == images == []


def test_unsupported_image_format(datadir):
    datadir.chdir()

    archive = ZipFile("sample_with_unsupported_image_format.xlsx", "r")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    images = find_images(archive, path)
    assert images == ([], [])


def test_find_images_accepts_sp_locks_no_text_edit():
    """Excel text boxes emit a:spLocks noTextEdit; that must not drop the drawing."""
    drawing = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <xdr:oneCellAnchor>
        <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
                   <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:ext cx="100" cy="100"/>
        <xdr:graphicFrame>
          <xdr:nvGraphicFramePr>
            <xdr:cNvPr id="1" name="Chart 1"/>
            <xdr:cNvGraphicFramePr/>
          </xdr:nvGraphicFramePr>
          <xdr:xfrm/>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       r:id="rId1"/>
            </a:graphicData>
          </a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData/>
      </xdr:oneCellAnchor>
      <xdr:twoCellAnchor>
        <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
                   <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
                 <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp macro="" textlink="$A$1">
          <xdr:nvSpPr>
            <xdr:cNvPr id="2" name="Text Box 1"/>
            <xdr:cNvSpPr txBox="1">
              <a:spLocks noChangeArrowheads="1" noTextEdit="1"/>
            </xdr:cNvSpPr>
          </xdr:nvSpPr>
          <xdr:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </xdr:spPr>
          <xdr:txBody>
            <a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>hello</a:t></a:r></a:p>
          </xdr:txBody>
        </xdr:sp>
        <xdr:clientData/>
      </xdr:twoCellAnchor>
    </xdr:wsDr>
    """
    buf = BytesIO()
    with ZipFile(buf, "w") as archive:
        archive.writestr("drawing1.xml", drawing)

    from fastpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
    from fastpyxl.xml.functions import fromstring
    from ..drawings import find_images

    parsed = SpreadsheetDrawing.from_tree(fromstring(drawing))
    assert len(parsed._chart_rels) == 1

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        charts, images = find_images(ZipFile(buf), "drawing1.xml")

    assert charts == images == []
    assert not any("DrawingML support is incomplete" in str(w.message) for w in caught)


def test_load_workbook_accepts_sp_locks_no_text_edit():
    """load_workbook must not warn/drop drawings solely due to noTextEdit locks."""
    drawing = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <xdr:twoCellAnchor>
        <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
                   <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
                 <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp macro="" textlink="$A$1">
          <xdr:nvSpPr>
            <xdr:cNvPr id="2" name="Text Box 1"/>
            <xdr:cNvSpPr txBox="1">
              <a:spLocks noChangeArrowheads="1" noTextEdit="1"/>
            </xdr:cNvSpPr>
          </xdr:nvSpPr>
          <xdr:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </xdr:spPr>
          <xdr:txBody>
            <a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>hello</a:t></a:r></a:p>
          </xdr:txBody>
        </xdr:sp>
        <xdr:clientData/>
      </xdr:twoCellAnchor>
    </xdr:wsDr>
    """
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/drawings/drawing1.xml"
        ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
    </Types>
    """
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        Target="xl/workbook.xml"/>
    </Relationships>
    """
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
    </workbook>
    """
    wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.xml"/>
    </Relationships>
    """
    sheet = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetData/>
      <drawing r:id="rId1"/>
    </worksheet>
    """
    sheet_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
        Target="../drawings/drawing1.xml"/>
    </Relationships>
    """
    buf = BytesIO()
    with ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels)
        z.writestr("xl/drawings/drawing1.xml", drawing)

    from fastpyxl import load_workbook

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        load_workbook(buf)

    assert not any("DrawingML support is incomplete" in str(w.message) for w in caught)
