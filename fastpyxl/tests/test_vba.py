# Copyright (c) 2010-2024 fastpyxl


# Python stdlib imports
from tempfile import NamedTemporaryFile
import zipfile

# package imports
from fastpyxl.reader.excel import load_workbook
from fastpyxl.xml.functions import fromstring
from fastpyxl.xml.constants import CONTYPES_NS


def test_content_types(datadir):
    datadir.join('reader').chdir()
    fname = 'vba+comments.xlsm'
    wb = load_workbook(fname, keep_vba=True)
    tmp = NamedTemporaryFile()
    wb.save(tmp)
    ct = fromstring(zipfile.ZipFile(tmp, 'r').open('[Content_Types].xml').read())
    s = set()
    for el in ct.findall("{%s}Override" % CONTYPES_NS):
        pn = el.get('PartName')
        assert pn not in s, 'duplicate PartName in [Content_Types].xml'
        s.add(pn)


def test_save_with_vba(datadir):
    datadir.join('reader').chdir()
    fname = 'vba-test.xlsm'
    wb = load_workbook(fname, keep_vba=True)
    tmp = NamedTemporaryFile()
    wb.save(tmp)
    files = set(zipfile.ZipFile(tmp, 'r').namelist())
    expected = set(['xl/drawings/_rels/vmlDrawing1.vml.rels',
                    'xl/worksheets/_rels/sheet1.xml.rels',
                    '[Content_Types].xml',
                    'xl/drawings/vmlDrawing1.vml',
                    'xl/ctrlProps/ctrlProp1.xml',
                    'xl/vbaProject.bin',
                    'docProps/core.xml',
                    '_rels/.rels',
                    'xl/theme/theme1.xml',
                    'xl/_rels/workbook.xml.rels',
                    'customUI/customUI.xml',
                    'xl/styles.xml',
                    'xl/worksheets/sheet1.xml',
                    'docProps/app.xml',
                    'xl/ctrlProps/ctrlProp2.xml',
                    'xl/workbook.xml',
                    'xl/activeX/activeX2.bin',
                    'xl/activeX/activeX1.bin',
                    'xl/media/image2.emf',
                    'xl/activeX/activeX1.xml',
                    'xl/activeX/_rels/activeX2.xml.rels',
                    'xl/media/image1.emf',
                    'xl/activeX/_rels/activeX1.xml.rels',
                    'xl/activeX/activeX2.xml',
                    ])
    assert files == expected


def test_save_with_saved_comments(datadir):
    datadir.join('reader').chdir()
    fname = 'vba-comments-saved.xlsm'
    wb = load_workbook(fname, keep_vba=True)
    tmp = NamedTemporaryFile()
    wb.save(tmp)
    files = set(zipfile.ZipFile(tmp, 'r').namelist())
    expected = set([
        'xl/styles.xml',
        'docProps/core.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/drawings/vmlDrawing1.vml',
        'xl/comments/comment1.xml',
        'docProps/app.xml',
        '[Content_Types].xml',
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/_rels/sheet1.xml.rels',
        '_rels/.rels',
        'xl/workbook.xml',
        'xl/theme/theme1.xml'
    ])
    assert files == expected

def test_vba_archive_excludes_workbook_files(datadir):
    """vba_archive should not store regular workbook XML files — only VBA-relevant ones."""
    datadir.join('reader').chdir()
    # These are large workbook files that must NOT be copied into vba_archive.
    NON_VBA = {'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/styles.xml',
               'xl/sharedStrings.xml', 'docProps/app.xml', 'docProps/core.xml'}
    wb = load_workbook('vba-test.xlsm', keep_vba=True)
    assert wb.vba_archive is not None
    vba_names = set(wb.vba_archive.namelist())
    overlap = vba_names & NON_VBA
    assert not overlap, f"Non-VBA workbook files found in vba_archive: {overlap}"
    # Sanity: vba_archive must be a strict subset of the original ZIP.
    with zipfile.ZipFile('vba-test.xlsm') as src:
        all_names = set(src.namelist())
    assert len(vba_names) < len(all_names), (
        "vba_archive should contain fewer files than the full archive"
    )


def test_vml_two_digit_index_in_vba_archive(datadir):
    """vmlDrawing10+ (two-digit indices) must be copied into vba_archive (issue #44)."""
    import io
    from fastpyxl.reader.excel import ARC_VBA
    # single-digit names must still match
    assert ARC_VBA.match('xl/drawings/vmlDrawing1.vml'), "single-digit should match"
    assert ARC_VBA.match('xl/drawings/vmlDrawing9.vml'), "single-digit should match"
    # two-digit (and higher) names must also match
    assert ARC_VBA.match('xl/drawings/vmlDrawing10.vml'), "two-digit should match"
    assert ARC_VBA.match('xl/drawings/vmlDrawing22.vml'), "two-digit should match"
    assert ARC_VBA.match('xl/drawings/vmlDrawing100.vml'), "three-digit should match"


def test_save_without_vba(datadir):
    datadir.join('reader').chdir()
    fname = 'vba-test.xlsm'
    vbFiles = set(['xl/activeX/activeX2.xml',
                   'xl/drawings/_rels/vmlDrawing1.vml.rels',
                   'xl/activeX/_rels/activeX1.xml.rels',
                   'xl/drawings/vmlDrawing1.vml',
                   'xl/activeX/activeX1.bin',
                   'xl/media/image1.emf',
                   'xl/vbaProject.bin',
                   'xl/activeX/_rels/activeX2.xml.rels',
                   'xl/worksheets/_rels/sheet1.xml.rels',
                   'customUI/customUI.xml',
                   'xl/media/image2.emf',
                   'xl/ctrlProps/ctrlProp1.xml',
                   'xl/activeX/activeX2.bin',
                   'xl/activeX/activeX1.xml',
                   'xl/ctrlProps/ctrlProp2.xml',
                   'xl/drawings/drawing1.xml'
                   ])

    wb = load_workbook(fname, keep_vba=False)
    tmp = NamedTemporaryFile()
    wb.save(tmp)
    files1 = set(zipfile.ZipFile(fname, 'r').namelist())
    files1.discard('xl/sharedStrings.xml')
    files2 = set(zipfile.ZipFile(tmp, 'r').namelist())
    difference = files1.difference(files2)
    assert difference.issubset(vbFiles), "Missing files: %s" % ', '.join(difference - vbFiles)
