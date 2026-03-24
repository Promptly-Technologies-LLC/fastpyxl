from fastpyxl.chartsheet.relation import DrawingHF, SheetBackgroundPicture
from fastpyxl.chartsheet.custom import CustomChartsheetView, CustomChartsheetViews
from fastpyxl.chartsheet.chartsheet import Chartsheet
from fastpyxl.chartsheet.protection import ChartsheetProtection
from fastpyxl.chartsheet.properties import ChartsheetProperties
from fastpyxl.chartsheet.publish import WebPublishItem, WebPublishItems
from fastpyxl.chartsheet.views import ChartsheetView, ChartsheetViewList
import datetime

from fastpyxl.packaging.core import DocumentProperties
from fastpyxl.packaging.relationship import Relationship
from fastpyxl.packaging.extended import ExtendedProperties
from fastpyxl.packaging.manifest import Manifest, Override
from fastpyxl.packaging.custom import CustomPropertyList, IntProperty, StringProperty
from fastpyxl.packaging.workbook import ChildSheet, FileRecoveryProperties, PivotCache
from fastpyxl.workbook.function_group import FunctionGroup, FunctionGroupList
from fastpyxl.workbook.properties import CalcProperties, FileVersion, WorkbookProperties
from fastpyxl.workbook.protection import FileSharing, WorkbookProtection
from fastpyxl.workbook.external_reference import ExternalReference
from fastpyxl.workbook.external_link.external import ExternalBook, ExternalDefinedName, ExternalLink, ExternalSheetNames
from fastpyxl.workbook.defined_name import DefinedName, DefinedNameList
from fastpyxl.workbook.smart_tags import SmartTag, SmartTagList, SmartTagProperties
from fastpyxl.workbook.views import BookView, CustomWorkbookView
from fastpyxl.workbook.web import WebPublishObject, WebPublishObjectList, WebPublishing
from fastpyxl.styles.colors import Color
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.functions import fromstring, tostring
from fastpyxl import __version__


def test_relationship_migrated_to_typed_base():
    assert issubclass(Relationship, TypedSerialisable)
    assert issubclass(ChartsheetProperties, TypedSerialisable)
    assert issubclass(ChartsheetView, TypedSerialisable)
    assert issubclass(ChartsheetViewList, TypedSerialisable)
    assert issubclass(WebPublishItem, TypedSerialisable)
    assert issubclass(WebPublishItems, TypedSerialisable)
    assert issubclass(SheetBackgroundPicture, TypedSerialisable)
    assert issubclass(DrawingHF, TypedSerialisable)
    assert issubclass(FunctionGroup, TypedSerialisable)
    assert issubclass(FunctionGroupList, TypedSerialisable)
    assert issubclass(WorkbookProperties, TypedSerialisable)
    assert issubclass(CalcProperties, TypedSerialisable)
    assert issubclass(FileVersion, TypedSerialisable)
    assert issubclass(BookView, TypedSerialisable)
    assert issubclass(CustomWorkbookView, TypedSerialisable)
    assert issubclass(WebPublishObject, TypedSerialisable)
    assert issubclass(WebPublishObjectList, TypedSerialisable)
    assert issubclass(WebPublishing, TypedSerialisable)
    assert issubclass(WorkbookProtection, TypedSerialisable)
    assert issubclass(FileSharing, TypedSerialisable)
    assert issubclass(ExternalReference, TypedSerialisable)
    assert issubclass(ExternalLink, TypedSerialisable)
    assert issubclass(ExternalBook, TypedSerialisable)
    assert issubclass(DefinedName, TypedSerialisable)
    assert issubclass(DefinedNameList, TypedSerialisable)
    assert issubclass(SmartTag, TypedSerialisable)
    assert issubclass(SmartTagList, TypedSerialisable)
    assert issubclass(SmartTagProperties, TypedSerialisable)
    assert issubclass(CustomChartsheetView, TypedSerialisable)
    assert issubclass(CustomChartsheetViews, TypedSerialisable)
    assert issubclass(ChartsheetProtection, TypedSerialisable)
    assert issubclass(Chartsheet, TypedSerialisable)
    assert issubclass(ExtendedProperties, TypedSerialisable)
    assert issubclass(Manifest, TypedSerialisable)
    assert issubclass(DocumentProperties, TypedSerialisable)
    assert issubclass(FileRecoveryProperties, TypedSerialisable)
    assert issubclass(ChildSheet, TypedSerialisable)
    assert issubclass(PivotCache, TypedSerialisable)


def test_relationship_alias_and_xml_roundtrip():
    model = Relationship(type="worksheet", Id="rId1", Target="worksheets/sheet1.xml")
    assert model.id == "rId1"
    assert model.target == "worksheets/sheet1.xml"

    xml = tostring(model.to_tree())
    expected = """
    <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" Id="rId1"></Relationship>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = Relationship.from_tree(fromstring(xml))
    assert parsed.Type == model.Type
    assert parsed.Target == model.Target
    assert parsed.Id == model.Id


def test_chartsheet_properties_and_views_serialization():
    props = ChartsheetProperties(published=True, codeName="Sheet1", tabColor=Color(rgb="FF0000"))
    xml = tostring(props.to_tree())
    expected = """
    <sheetPr published="1" codeName="Sheet1">
      <tabColor rgb="00FF0000"></tabColor>
    </sheetPr>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    views = ChartsheetViewList(sheetView=[ChartsheetView(workbookViewId=0, zoomToFit=True)])
    view_xml = tostring(views.to_tree())
    view_expected = """
    <sheetViews>
      <sheetView workbookViewId="0" zoomToFit="1"></sheetView>
    </sheetViews>
    """
    diff = compare_xml(view_xml, view_expected)
    assert diff is None, diff


def test_web_publish_items_computed_count_serialization():
    items = WebPublishItems(webPublishItem=[WebPublishItem(id=1, divId="d1", sourceType="sheet", sourceRef="A1", destinationFile="out.htm")])
    xml = tostring(items.to_tree())
    expected = """
    <WebPublishItems count="1">
      <webPublishItem id="1" divId="d1" sourceType="sheet" sourceRef="A1" destinationFile="out.htm"></webPublishItem>
    </WebPublishItems>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_chartsheet_relation_alias_fields_serialize():
    model = DrawingHF(id="rId1", lho=1, cff=2)
    assert model.leftHeaderOddPages == 1
    assert model.centerFooterFirstPage == 2
    xml = tostring(model.to_tree("drawingHF"))
    expected = """
    <drawingHF xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" lho="1" cff="2" r:id="rId1"></drawingHF>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_function_group_list_serialization():
    model = FunctionGroupList(builtInGroupCount=16, functionGroup=[FunctionGroup(name="Cube")])
    xml = tostring(model.to_tree())
    expected = """
    <functionGroups builtInGroupCount="16">
      <functionGroup name="Cube"></functionGroup>
    </functionGroups>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_workbook_properties_and_views_serialization():
    props = WorkbookProperties(showObjects="all", updateLinks="never")
    props_xml = tostring(props.to_tree())
    props_expected = """
    <workbookPr showObjects="all" updateLinks="never"></workbookPr>
    """
    diff = compare_xml(props_xml, props_expected)
    assert diff is None, diff

    calc = CalcProperties(calcId=124519, fullCalcOnLoad=True)
    calc_xml = tostring(calc.to_tree())
    calc_expected = """
    <calcPr calcId="124519" fullCalcOnLoad="1"></calcPr>
    """
    diff = compare_xml(calc_xml, calc_expected)
    assert diff is None, diff

    view = BookView(activeTab=1)
    view_xml = tostring(view.to_tree())
    view_expected = """
    <workbookView visibility="visible" minimized="0" showHorizontalScroll="1" showVerticalScroll="1" showSheetTabs="1" tabRatio="600" firstSheet="0" activeTab="1" autoFilterDateGrouping="1"></workbookView>
    """
    diff = compare_xml(view_xml, view_expected)
    assert diff is None, diff


def test_custom_chartsheet_views_serialization():
    model = CustomChartsheetViews(
        customSheetView=[CustomChartsheetView(guid="{00000000-0000-0000-0000-000000000000}", scale=100, state="visible")]
    )
    xml = tostring(model.to_tree())
    expected = """
    <customSheetViews>
      <customSheetView guid="{00000000-0000-0000-0000-000000000000}" scale="100" state="visible"></customSheetView>
    </customSheetViews>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_extended_properties_serialization_includes_namespace():
    model = ExtendedProperties(
        Application="Microsoft Excel Compatible / Openpyxl",
        AppVersion="3.1",
        DocSecurity=0,
        ScaleCrop=False,
    )
    xml = tostring(model.to_tree())
    expected = """
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
      <Application>Microsoft Excel Compatible / Openpyxl</Application>
      <AppVersion>3.1</AppVersion>
      <DocSecurity>0</DocSecurity>
      <ScaleCrop>true</ScaleCrop>
    </Properties>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_manifest_preserves_unique_defaults_and_overrides():
    manifest = Manifest()
    manifest._append_default_if_missing(manifest.Default[0])
    manifest._append_override_if_missing(manifest.Override[0])
    manifest.append(type("Obj", (), {"path": "/xl/workbook.xml", "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"})())
    manifest.append(type("Obj", (), {"path": "/xl/workbook.xml", "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"})())

    assert len([d for d in manifest.Default if d.Extension == "rels"]) == 1
    assert len([o for o in manifest.Override if o.PartName == "/xl/workbook.xml"]) == 1

    xml = tostring(manifest.to_tree())
    expected = """
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"></Default>
      <Default Extension="xml" ContentType="application/xml"></Default>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"></Override>
      <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"></Override>
      <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"></Override>
      <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"></Override>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"></Override>
    </Types>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_packaging_workbook_leaf_models_serialize():
    recovery = FileRecoveryProperties(autoRecover=True, crashSave=False)
    recovery_xml = tostring(recovery.to_tree())
    recovery_expected = """
    <fileRecoveryPr autoRecover="1" crashSave="0"></fileRecoveryPr>
    """
    diff = compare_xml(recovery_xml, recovery_expected)
    assert diff is None, diff

    sheet = ChildSheet(name="Sheet1", sheetId=1, state="visible", id="rId1")
    sheet_xml = tostring(sheet.to_tree())
    sheet_expected = """
    <sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Sheet1" sheetId="1" state="visible" r:id="rId1"></sheet>
    """
    diff = compare_xml(sheet_xml, sheet_expected)
    assert diff is None, diff

    cache = PivotCache(cacheId=1, id="rId2")
    cache_xml = tostring(cache.to_tree())
    cache_expected = """
    <pivotCache xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" cacheId="1" r:id="rId2"></pivotCache>
    """
    diff = compare_xml(cache_xml, cache_expected)
    assert diff is None, diff


def test_document_properties_qualified_datetime_rendering():
    ts = datetime.datetime(2026, 3, 1, 2, 3, 4)
    model = DocumentProperties(
        creator="fastpyxl",
        created=ts,
        modified=ts,
        title="Book",
    )
    xml = tostring(model.to_tree())
    expected = """
    <coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
      <dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">fastpyxl</dc:creator>
      <dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">Book</dc:title>
      <dcterms:created xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">2026-03-01T02:03:04Z</dcterms:created>
      <dcterms:modified xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">2026-03-01T02:03:04Z</dcterms:modified>
    </coreProperties>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_workbook_web_models_serialization():
    obj_list = WebPublishObjectList(
        webPublishObject=[WebPublishObject(id=1, divId="d1", destinationFile="out.htm")]
    )
    list_xml = tostring(obj_list.to_tree())
    list_expected = """
    <webPublishingObjects count="1">
      <webPublishingObject id="1" divId="d1" destinationFile="out.htm"></webPublishingObject>
    </webPublishingObjects>
    """
    diff = compare_xml(list_xml, list_expected)
    assert diff is None, diff

    web = WebPublishing(targetScreenSize="800x600", allowPng=True)
    web_xml = tostring(web.to_tree())
    web_expected = """
    <webPublishing targetScreenSize="800x600" allowPng="1"></webPublishing>
    """
    diff = compare_xml(web_xml, web_expected)
    assert diff is None, diff


def test_workbook_protection_password_hashing_and_external_reference():
    protection = WorkbookProtection(workbookPassword="secret", lockStructure=True)
    assert protection.workbookPassword is not None
    assert protection.workbookPassword != "secret"
    xml = tostring(protection.to_tree())
    assert b'lockStructure="1"' in xml
    assert b'workbookPassword="' in xml

    sharing = FileSharing(readOnlyRecommended=True, userName="alice")
    sharing_xml = tostring(sharing.to_tree())
    sharing_expected = """
    <fileSharing readOnlyRecommended="1" userName="alice"></fileSharing>
    """
    diff = compare_xml(sharing_xml, sharing_expected)
    assert diff is None, diff

    external = ExternalReference(id="rId1")
    external_xml = tostring(external.to_tree())
    external_expected = """
    <externalReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"></externalReference>
    """
    diff = compare_xml(external_xml, external_expected)
    assert diff is None, diff


def test_workbook_smart_tags_serialization():
    tags = SmartTagList(smartTagType=[SmartTag(namespaceUri="urn:test", name="tag", url="https://x")])
    tags_xml = tostring(tags.to_tree())
    tags_expected = """
    <smartTagTypes>
      <smartTagType namespaceUri="urn:test" name="tag" url="https://x"></smartTagType>
    </smartTagTypes>
    """
    diff = compare_xml(tags_xml, tags_expected)
    assert diff is None, diff

    props = SmartTagProperties(embed=True, show="all")
    props_xml = tostring(props.to_tree())
    props_expected = """
    <smartTagPr embed="1" show="all"></smartTagPr>
    """
    diff = compare_xml(props_xml, props_expected)
    assert diff is None, diff


def test_chartsheet_and_protection_serialization():
    protection = ChartsheetProtection(content=True, objects=False)
    protection_xml = tostring(protection.to_tree())
    protection_expected = """
    <sheetProtection content="1" objects="0"></sheetProtection>
    """
    diff = compare_xml(protection_xml, protection_expected)
    assert diff is None, diff

    chartsheet = Chartsheet(
        title="Chart1",
        sheetPr=ChartsheetProperties(codeName="ChartSheet"),
        sheetProtection=protection,
    )
    chart_xml = tostring(chartsheet.to_tree())
    chart_expected = """
    <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetPr codeName="ChartSheet"></sheetPr>
      <sheetViews><sheetView workbookViewId="0" zoomToFit="1"></sheetView></sheetViews>
      <sheetProtection content="1" objects="0"></sheetProtection>
      <drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"></drawing>
    </chartsheet>
    """
    diff = compare_xml(chart_xml, chart_expected)
    assert diff is None, diff


def test_external_link_serialization():
    model = ExternalLink(
        externalBook=ExternalBook(
            sheetNames=ExternalSheetNames(sheetName=["Sheet1"]),
            definedNames=[ExternalDefinedName(name="N1", refersTo="'Sheet1'!$A$1", sheetId=0)],
            id="rId1",
        )
    )
    xml = tostring(model.to_tree())
    expected = """
    <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">
        <sheetNames>
          <sheetName val="Sheet1"></sheetName>
        </sheetNames>
        <definedNames>
          <definedName name="N1" refersTo="'Sheet1'!$A$1" sheetId="0"></definedName>
        </definedNames>
      </externalBook>
    </externalLink>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_custom_property_list_roundtrip_string_and_int():
    props = CustomPropertyList()
    props.append(StringProperty(name="Client", value="Acme"))
    props.append(IntProperty(name="Build", value=42))
    tree = props.to_tree()
    xml = tostring(tree)
    expected = """
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties">
      <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client">
        <vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">Acme</vt:lpwstr>
      </property>
      <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Build">
        <vt:i4 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">42</vt:i4>
      </property>
    </Properties>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_defined_name_serialization_and_destinations():
    dn = DefinedName(name="MyRange", attr_text="'Sheet1'!$A$1:$B$2", localSheetId=0, hidden=False)
    xml = tostring(dn.to_tree())
    expected = """
    <definedName name="MyRange" localSheetId="0" hidden="0">Sheet1!$A$1:$B$2</definedName>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    assert list(dn.destinations) == [("Sheet1", "$A$1:$B$2")]

    dnl = DefinedNameList(definedName=[dn])
    list_xml = tostring(dnl.to_tree())
    list_expected = """
    <definedNames>
      <definedName name="MyRange" localSheetId="0" hidden="0">Sheet1!$A$1:$B$2</definedName>
    </definedNames>
    """
    diff = compare_xml(list_xml, list_expected)
    assert diff is None, diff
