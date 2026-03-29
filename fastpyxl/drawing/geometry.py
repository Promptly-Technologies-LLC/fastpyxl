# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from .line import LineProperties  # noqa: F401 -- re-exported

from fastpyxl.styles.colors import Color
from fastpyxl.xml.constants import DRAWING_NS
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field


def _emu_coord(value, field_name: str):
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


def _range_converter(value, *, field_name: str, min_val, max_val):
    if value is None:
        return None
    try:
        numeric = int(value)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc
    if numeric < min_val or numeric > max_val:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return numeric


class Point2D(Serialisable):

    tagname = "off"
    namespace = DRAWING_NS

    x: int | None = Field.attribute(
        expected_type=int,
        allow_none=True,
        converter=lambda v: _emu_coord(v, "x"), default=None,
    )
    y: int | None = Field.attribute(
        expected_type=int,
        allow_none=True,
        converter=lambda v: _emu_coord(v, "y"), default=None,
    )

    def __init__(self,
                 x=None,
                 y=None,
                ):
        self.x = x
        self.y = y


class PositiveSize2D(Serialisable):

    tagname = "ext"
    namespace = DRAWING_NS

    """
    Dimensions in EMUs
    """

    cx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    width = AliasField("cx", default=None)
    cy: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    height = AliasField("cy", default=None)

    def __init__(self,
                 cx=None,
                 cy=None,
                ):
        self.cx = cx
        self.cy = cy


class Transform2D(Serialisable):

    tagname = "xfrm"
    namespace = DRAWING_NS

    rot: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    flipH: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    flipV: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    off: Point2D | None = Field.element(expected_type=Point2D, allow_none=True, default=None)
    ext: PositiveSize2D | None = Field.element(expected_type=PositiveSize2D, allow_none=True, default=None)
    chOff: Point2D | None = Field.element(expected_type=Point2D, allow_none=True, default=None)
    chExt: PositiveSize2D | None = Field.element(expected_type=PositiveSize2D, allow_none=True, default=None)

    xml_order = ("off", "ext", "chOff", "chExt")

    def __init__(self,
                 rot=None,
                 flipH=None,
                 flipV=None,
                 off=None,
                 ext=None,
                 chOff=None,
                 chExt=None,
                ):
        self.rot = rot
        self.flipH = flipH
        self.flipV = flipV
        self.off = off
        self.ext = ext
        self.chOff = chOff
        self.chExt = chExt


class GroupTransform2D(Serialisable):

    tagname = "xfrm"
    namespace = DRAWING_NS

    rot: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    flipH: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    flipV: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    off: Point2D | None = Field.element(expected_type=Point2D, allow_none=True, default=None)
    ext: PositiveSize2D | None = Field.element(expected_type=PositiveSize2D, allow_none=True, default=None)
    chOff: Point2D | None = Field.element(expected_type=Point2D, allow_none=True, default=None)
    chExt: PositiveSize2D | None = Field.element(expected_type=PositiveSize2D, allow_none=True, default=None)

    xml_order = ("off", "ext", "chOff", "chExt")

    def __init__(self,
                 rot=0,
                 flipH=None,
                 flipV=None,
                 off=None,
                 ext=None,
                 chOff=None,
                 chExt=None,
                ):
        self.rot = rot
        self.flipH = flipH
        self.flipV = flipV
        self.off = off
        self.ext = ext
        self.chOff = chOff
        self.chExt = chExt


class SphereCoords(Serialisable):

    tagname = "sphereCoords"

    lat: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    lon: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rev: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 lat=None,
                 lon=None,
                 rev=None,
                ):
        self.lat = lat
        self.lon = lon
        self.rev = rev


_CAMERA_PRST_VALUES = frozenset([
    'legacyObliqueTopLeft', 'legacyObliqueTop', 'legacyObliqueTopRight', 'legacyObliqueLeft',
    'legacyObliqueFront', 'legacyObliqueRight', 'legacyObliqueBottomLeft',
    'legacyObliqueBottom', 'legacyObliqueBottomRight', 'legacyPerspectiveTopLeft',
    'legacyPerspectiveTop', 'legacyPerspectiveTopRight', 'legacyPerspectiveLeft',
    'legacyPerspectiveFront', 'legacyPerspectiveRight', 'legacyPerspectiveBottomLeft',
    'legacyPerspectiveBottom', 'legacyPerspectiveBottomRight', 'orthographicFront',
    'isometricTopUp', 'isometricTopDown', 'isometricBottomUp', 'isometricBottomDown',
    'isometricLeftUp', 'isometricLeftDown', 'isometricRightUp', 'isometricRightDown',
    'isometricOffAxis1Left', 'isometricOffAxis1Right', 'isometricOffAxis1Top',
    'isometricOffAxis2Left', 'isometricOffAxis2Right', 'isometricOffAxis2Top',
    'isometricOffAxis3Left', 'isometricOffAxis3Right', 'isometricOffAxis3Bottom',
    'isometricOffAxis4Left', 'isometricOffAxis4Right', 'isometricOffAxis4Bottom',
    'obliqueTopLeft', 'obliqueTop', 'obliqueTopRight', 'obliqueLeft', 'obliqueRight',
    'obliqueBottomLeft', 'obliqueBottom', 'obliqueBottomRight', 'perspectiveFront',
    'perspectiveLeft', 'perspectiveRight', 'perspectiveAbove', 'perspectiveBelow',
    'perspectiveAboveLeftFacing', 'perspectiveAboveRightFacing',
    'perspectiveContrastingLeftFacing', 'perspectiveContrastingRightFacing',
    'perspectiveHeroicLeftFacing', 'perspectiveHeroicRightFacing',
    'perspectiveHeroicExtremeLeftFacing', 'perspectiveHeroicExtremeRightFacing',
    'perspectiveRelaxed', 'perspectiveRelaxedModerately',
])

_LIGHT_RIG_VALUES = frozenset([
    'legacyFlat1', 'legacyFlat2', 'legacyFlat3', 'legacyFlat4', 'legacyNormal1',
    'legacyNormal2', 'legacyNormal3', 'legacyNormal4', 'legacyHarsh1',
    'legacyHarsh2', 'legacyHarsh3', 'legacyHarsh4', 'threePt', 'balanced',
    'soft', 'harsh', 'flood', 'contrasting', 'morning', 'sunrise', 'sunset',
    'chilly', 'freezing', 'flat', 'twoPt', 'glow', 'brightRoom',
])

_DIR_VALUES = frozenset(['tl', 't', 'tr', 'l', 'r', 'bl', 'b', 'br'])

_BEVEL_VALUES = frozenset([
    'relaxedInset', 'circle', 'slope', 'cross', 'angle',
    'softRound', 'convex', 'coolSlant', 'divot', 'riblet',
    'hardEdge', 'artDeco',
])

_PRST_MATERIAL_VALUES = frozenset([
    'legacyMatte', 'legacyPlastic', 'legacyMetal', 'legacyWireframe', 'matte', 'plastic',
    'metal', 'warmMatte', 'translucentPowder', 'powder', 'dkEdge',
    'softEdge', 'clear', 'flat', 'softmetal',
])

_FILL_VALUES = frozenset(['norm', 'lighten', 'lightenLess', 'darken', 'darkenLess'])

_FONT_REF_VALUES = frozenset(['major', 'minor'])

_PRST_GEOM_VALUES = frozenset([
    'line', 'lineInv', 'triangle', 'rtTriangle', 'rect',
    'diamond', 'parallelogram', 'trapezoid', 'nonIsoscelesTrapezoid',
    'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon',
    'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12',
    'star16', 'star24', 'star32', 'roundRect', 'round1Rect',
    'round2SameRect', 'round2DiagRect', 'snipRoundRect', 'snip1Rect',
    'snip2SameRect', 'snip2DiagRect', 'plaque', 'ellipse', 'teardrop',
    'homePlate', 'chevron', 'pieWedge', 'pie', 'blockArc', 'donut',
    'noSmoking', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
    'stripedRightArrow', 'notchedRightArrow', 'bentUpArrow',
    'leftRightArrow', 'upDownArrow', 'leftUpArrow', 'leftRightUpArrow',
    'quadArrow', 'leftArrowCallout', 'rightArrowCallout', 'upArrowCallout',
    'downArrowCallout', 'leftRightArrowCallout', 'upDownArrowCallout',
    'quadArrowCallout', 'bentArrow', 'uturnArrow', 'circularArrow',
    'leftCircularArrow', 'leftRightCircularArrow', 'curvedRightArrow',
    'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'swooshArrow',
    'cube', 'can', 'lightningBolt', 'heart', 'sun', 'moon', 'smileyFace',
    'irregularSeal1', 'irregularSeal2', 'foldedCorner', 'bevel', 'frame',
    'halfFrame', 'corner', 'diagStripe', 'chord', 'arc', 'leftBracket',
    'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair', 'bracePair',
    'straightConnector1', 'bentConnector2', 'bentConnector3',
    'bentConnector4', 'bentConnector5', 'curvedConnector2',
    'curvedConnector3', 'curvedConnector4', 'curvedConnector5', 'callout1',
    'callout2', 'callout3', 'accentCallout1', 'accentCallout2',
    'accentCallout3', 'borderCallout1', 'borderCallout2', 'borderCallout3',
    'accentBorderCallout1', 'accentBorderCallout2', 'accentBorderCallout3',
    'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout',
    'cloudCallout', 'cloud', 'ribbon', 'ribbon2', 'ellipseRibbon',
    'ellipseRibbon2', 'leftRightRibbon', 'verticalScroll',
    'horizontalScroll', 'wave', 'doubleWave', 'plus', 'flowChartProcess',
    'flowChartDecision', 'flowChartInputOutput',
    'flowChartPredefinedProcess', 'flowChartInternalStorage',
    'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator',
    'flowChartPreparation', 'flowChartManualInput',
    'flowChartManualOperation', 'flowChartConnector', 'flowChartPunchedCard',
    'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartOr',
    'flowChartCollate', 'flowChartSort', 'flowChartExtract',
    'flowChartMerge', 'flowChartOfflineStorage', 'flowChartOnlineStorage',
    'flowChartMagneticTape', 'flowChartMagneticDisk',
    'flowChartMagneticDrum', 'flowChartDisplay', 'flowChartDelay',
    'flowChartAlternateProcess', 'flowChartOffpageConnector',
    'actionButtonBlank', 'actionButtonHome', 'actionButtonHelp',
    'actionButtonInformation', 'actionButtonForwardNext',
    'actionButtonBackPrevious', 'actionButtonEnd', 'actionButtonBeginning',
    'actionButtonReturn', 'actionButtonDocument', 'actionButtonSound',
    'actionButtonMovie', 'gear6', 'gear9', 'funnel', 'mathPlus', 'mathMinus',
    'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual', 'cornerTabs',
    'squareTabs', 'plaqueTabs', 'chartX', 'chartStar', 'chartPlus',
])


class Camera(Serialisable):

    tagname = "camera"

    prst: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _CAMERA_PRST_VALUES, "prst"), default=None,
    )
    fov: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    zoom: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rot: SphereCoords | None = Field.element(expected_type=SphereCoords, allow_none=True, default=None)

    def __init__(self,
                 prst=None,
                 fov=None,
                 zoom=None,
                 rot=None,
                ):
        self.prst = prst
        self.fov = fov
        self.zoom = zoom
        self.rot = rot


class LightRig(Serialisable):

    tagname = "lightRig"

    rig: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _LIGHT_RIG_VALUES, "rig"), default=None,
    )
    dir: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _DIR_VALUES, "dir"), default=None,
    )
    rot: SphereCoords | None = Field.element(expected_type=SphereCoords, allow_none=True, default=None)

    def __init__(self,
                 rig=None,
                 dir=None,
                 rot=None,
                ):
        self.rig = rig
        self.dir = dir
        self.rot = rot


class Vector3D(Serialisable):

    tagname = "vector"

    dx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    dy: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    dz: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 dx=None,
                 dy=None,
                 dz=None,
                ):
        self.dx = dx
        self.dy = dy
        self.dz = dz


class Point3D(Serialisable):

    tagname = "anchor"

    x: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    y: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    z: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 x=None,
                 y=None,
                 z=None,
                ):
        self.x = x
        self.y = y
        self.z = z


class Backdrop(Serialisable):

    tagname = "backdrop"

    anchor: Point3D | None = Field.element(expected_type=Point3D, allow_none=True, default=None)
    norm: Vector3D | None = Field.element(expected_type=Vector3D, allow_none=True, default=None)
    up: Vector3D | None = Field.element(expected_type=Vector3D, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 anchor=None,
                 norm=None,
                 up=None,
                 extLst=None,
                ):
        self.anchor = anchor
        self.norm = norm
        self.up = up
        self.extLst = extLst


class Scene3D(Serialisable):

    tagname = "scene3d"

    camera: Camera | None = Field.element(expected_type=Camera, allow_none=True, default=None)
    lightRig: LightRig | None = Field.element(expected_type=LightRig, allow_none=True, default=None)
    backdrop: Backdrop | None = Field.element(expected_type=Backdrop, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 camera=None,
                 lightRig=None,
                 backdrop=None,
                 extLst=None,
                ):
        self.camera = camera
        self.lightRig = lightRig
        self.backdrop = backdrop
        self.extLst = extLst


class Bevel(Serialisable):

    tagname = "bevel"

    w: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    h: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    prst: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _BEVEL_VALUES, "prst"), default=None,
    )

    def __init__(self,
                 w=None,
                 h=None,
                 prst=None,
                ):
        self.w = w
        self.h = h
        self.prst = prst


class Shape3D(Serialisable):

    tagname = "sp3d"
    namespace = DRAWING_NS

    z: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    extrusionH: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    contourW: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    prstMaterial: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _PRST_MATERIAL_VALUES, "prstMaterial"), default=None,
    )
    bevelT: Bevel | None = Field.element(expected_type=Bevel, allow_none=True, default=None)
    bevelB: Bevel | None = Field.element(expected_type=Bevel, allow_none=True, default=None)
    extrusionClr: Color | None = Field.element(expected_type=Color, allow_none=True, default=None)
    contourClr: Color | None = Field.element(expected_type=Color, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 z=None,
                 extrusionH=None,
                 contourW=None,
                 prstMaterial=None,
                 bevelT=None,
                 bevelB=None,
                 extrusionClr=None,
                 contourClr=None,
                 extLst=None,
                ):
        self.z = z
        self.extrusionH = extrusionH
        self.contourW = contourW
        self.prstMaterial = prstMaterial
        self.bevelT = bevelT
        self.bevelB = bevelB
        self.extrusionClr = extrusionClr
        self.contourClr = contourClr
        self.extLst = extLst


class Path2D(Serialisable):

    tagname = "path"

    w: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    h: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    fill: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _FILL_VALUES, "fill"), default=None,
    )
    stroke: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    extrusionOk: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self,
                 w=None,
                 h=None,
                 fill=None,
                 stroke=None,
                 extrusionOk=None,
                ):
        self.w = w
        self.h = h
        self.fill = fill
        self.stroke = stroke
        self.extrusionOk = extrusionOk


class Path2DList(Serialisable):

    tagname = "pathLst"

    path: Path2D | None = Field.element(expected_type=Path2D, allow_none=True, default=None)

    def __init__(self,
                 path=None,
                ):
        self.path = path


class GeomRect(Serialisable):

    tagname = "rect"

    l: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)  # noqa: E741
    t: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    r: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    b: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(
                 self,
                 l=None,  # noqa: E741
                 t=None,
                 r=None,
                 b=None,
                ):
        self.l = l  # noqa: E741
        self.t = t
        self.r = r
        self.b = b


class AdjPoint2D(Serialisable):

    tagname = "pt"

    x: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    y: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 x=None,
                 y=None,
                ):
        self.x = x
        self.y = y


class ConnectionSite(Serialisable):

    tagname = "cxn"

    ang: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    pos: AdjPoint2D | None = Field.element(expected_type=AdjPoint2D, allow_none=True, default=None)

    def __init__(self,
                 ang=None,
                 pos=None,
                ):
        self.ang = ang
        self.pos = pos


class ConnectionSiteList(Serialisable):

    tagname = "cxnLst"

    cxn: ConnectionSite | None = Field.element(expected_type=ConnectionSite, allow_none=True, default=None)

    def __init__(self,
                 cxn=None,
                ):
        self.cxn = cxn


class AdjustHandleList(Serialisable):

    tagname = "ahLst"


class GeomGuide(Serialisable):

    tagname = "gd"

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fmla: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 name=None,
                 fmla=None,
                ):
        self.name = name
        self.fmla = fmla


class GeomGuideList(Serialisable):

    tagname = "gdLst"

    gd: GeomGuide | None = Field.element(expected_type=GeomGuide, allow_none=True, default=None)

    def __init__(self,
                 gd=None,
                ):
        self.gd = gd


class CustomGeometry2D(Serialisable):

    tagname = "custGeom"

    avLst: GeomGuideList | None = Field.element(expected_type=GeomGuideList, allow_none=True, default=None)
    gdLst: GeomGuideList | None = Field.element(expected_type=GeomGuideList, allow_none=True, default=None)
    ahLst: AdjustHandleList | None = Field.element(expected_type=AdjustHandleList, allow_none=True, default=None)
    cxnLst: ConnectionSiteList | None = Field.element(expected_type=ConnectionSiteList, allow_none=True, default=None)
    pathLst: Path2DList | None = Field.element(expected_type=Path2DList, allow_none=True, default=None)

    def __init__(self,
                 avLst=None,
                 gdLst=None,
                 ahLst=None,
                 cxnLst=None,
                 rect=None,
                 pathLst=None,
                ):
        self.avLst = avLst
        self.gdLst = gdLst
        self.ahLst = ahLst
        self.cxnLst = cxnLst
        self.rect = None
        self.pathLst = pathLst


class PresetGeometry2D(Serialisable):

    tagname = "prstGeom"
    namespace = DRAWING_NS

    prst: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _PRST_GEOM_VALUES, "prst"), default=None,
    )
    avLst: GeomGuideList | None = Field.element(expected_type=GeomGuideList, allow_none=True, default=None)

    def __init__(self,
                 prst=None,
                 avLst=None,
                ):
        self.prst = prst
        self.avLst = avLst


class FontReference(Serialisable):

    tagname = "fontRef"

    idx: str | None = Field.attribute(
        expected_type=str, allow_none=True,
        converter=lambda v: _enum_converter(v, _FONT_REF_VALUES, "idx"), default=None,
    )

    def __init__(self,
                 idx=None,
                ):
        self.idx = idx


class StyleMatrixReference(Serialisable):

    tagname = "styleRef"

    idx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 idx=None,
                ):
        self.idx = idx


class ShapeStyle(Serialisable):

    tagname = "style"

    lnRef: StyleMatrixReference | None = Field.element(expected_type=StyleMatrixReference, allow_none=True, default=None)
    fillRef: StyleMatrixReference | None = Field.element(expected_type=StyleMatrixReference, allow_none=True, default=None)
    effectRef: StyleMatrixReference | None = Field.element(expected_type=StyleMatrixReference, allow_none=True, default=None)
    fontRef: FontReference | None = Field.element(expected_type=FontReference, allow_none=True, default=None)

    def __init__(self,
                 lnRef=None,
                 fillRef=None,
                 effectRef=None,
                 fontRef=None,
                ):
        self.lnRef = lnRef
        self.fillRef = fillRef
        self.effectRef = effectRef
        self.fontRef = fontRef
