Working with Images
===================

Inserting an image
-------------------
.. :: doctest

>>> from fastpyxl import Workbook
>>> from fastpyxl.drawing.image import Image
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>> ws['A1'] = 'You should see three logos below'
>>>
>>> # create an image
>>> img = Image('logo.png')
>>>
>>> # add to worksheet and anchor next to cells
>>> ws.add_image(img, 'A1')
>>> wb.save('logo.xlsx')