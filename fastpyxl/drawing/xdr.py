# Copyright (c) 2010-2024 fastpyxl

"""
Spreadsheet Drawing has some copies of Drawing ML elements
"""

from .geometry import Point2D, PositiveSize2D, Transform2D


class XDRPoint2D(Point2D):

    namespace = None


class XDRPositiveSize2D(PositiveSize2D):

    namespace = None


class XDRTransform2D(Transform2D):

    namespace = None
