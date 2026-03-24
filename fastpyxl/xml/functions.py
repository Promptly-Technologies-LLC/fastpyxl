# Copyright (c) 2010-2024 fastpyxl

"""
XML compatibility functions
"""

# Python stdlib imports
import re

from fastpyxl import DEFUSEDXML, LXML

if LXML is True:
    from lxml.etree import (  # noqa: F401
    Element,
    SubElement,
    register_namespace,
    QName,
    xmlfile,
    XMLParser,
    )
    from lxml.etree import XMLSyntaxError
    from lxml.etree import fromstring as _lxml_fromstring, tostring
    # Do not load external DTDs/entities; keep parsing failures as ValueError for callers.
    safe_parser = XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        huge_tree=False,
    )

    def fromstring(source, parser=None):
        from_file = hasattr(source, "read")
        if from_file:
            source = source.read()
        if from_file and isinstance(source, (bytes, bytearray)) and b"<!doctype" in source.lower():
            raise ValueError("DOCTYPE declarations are not supported")
        if from_file and isinstance(source, str) and "<!doctype" in source.lower():
            raise ValueError("DOCTYPE declarations are not supported")
        try:
            return _lxml_fromstring(source, parser=parser or safe_parser)
        except XMLSyntaxError as exc:
            raise ValueError(str(exc)) from exc

else:
    from xml.etree.ElementTree import (  # noqa: F401
    Element,
    SubElement,
    fromstring as _stdlib_fromstring,
    tostring,
    QName,
    register_namespace
    )
    from et_xmlfile import xmlfile  # noqa: F401
    if DEFUSEDXML is True:
        from defusedxml.ElementTree import fromstring as _stdlib_fromstring

    def fromstring(source, parser=None):
        if hasattr(source, "read"):
            source = source.read()
        return _stdlib_fromstring(source)

from xml.etree.ElementTree import iterparse  # noqa: F401

from fastpyxl.xml.constants import (
    CHART_NS,
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_DRAWING_NS,
    SHEET_MAIN_NS,
    REL_NS,
    VTYPES_NS,
    COREPROPS_NS,
    CUSTPROPS_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX,
    XML_NS
)

register_namespace(DCTERMS_PREFIX, DCTERMS_NS)
register_namespace('dcmitype', 'http://purl.org/dc/dcmitype/')
register_namespace('cp', COREPROPS_NS)
register_namespace('c', CHART_NS)
register_namespace('a', DRAWING_NS)
register_namespace('s', SHEET_MAIN_NS)
register_namespace('r', REL_NS)
register_namespace('vt', VTYPES_NS)
register_namespace('xdr', SHEET_DRAWING_NS)
register_namespace('cdr', CHART_DRAWING_NS)
register_namespace('xml', XML_NS)
register_namespace('cust', CUSTPROPS_NS)


from functools import partial

tostring = partial(tostring, encoding="utf-8")

NS_REGEX = re.compile("({(?P<namespace>.*)})?(?P<localname>.*)")

def localname(node):
    if callable(node.tag):
        return "comment"
    m = NS_REGEX.match(node.tag)
    if m is None:
        return node.tag
    return m.group('localname')


def whitespace(node):
    stripped = node.text.strip()
    if stripped and node.text != stripped:
        node.set("{%s}space" % XML_NS, "preserve")
