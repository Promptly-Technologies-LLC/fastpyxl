Introduction
------------

fastpyxl is a high-performance Python library to read/write Excel 2010
xlsx/xlsm/xltx/xltm files. It is a fork of `openpyxl <https://foss.heptapod.net/openpyxl/openpyxl>`_,
focused on optimized performance for large-scale Excel processing.

All kudos to the PHPExcel team and the openpyxl developers whose work this
project builds upon.


Security
--------

By default fastpyxl does not guard against quadratic blowup or billion laughs
xml attacks. To guard against these attacks install defusedxml.


Sample code::

    from fastpyxl import Workbook
    wb = Workbook()

    # grab the active worksheet
    ws = wb.active

    # Data can be assigned directly to cells
    ws['A1'] = 42

    # Rows can also be appended
    ws.append([1, 2, 3])

    # Python types will automatically be converted
    import datetime
    ws['A2'] = datetime.datetime.now()

    # Save the file
    wb.save("sample.xlsx")


Documentation
-------------

The documentation is at: https://fastpyxl.readthedocs.io

* installation methods
* code examples
* instructions for contributing

Release notes: https://fastpyxl.readthedocs.io/en/stable/changes.html
