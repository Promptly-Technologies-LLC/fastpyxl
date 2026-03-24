# Copyright (c) 2010-2024 fastpyxl

DEBUG = False

from fastpyxl.compat.numbers import NUMPY
from fastpyxl.xml import DEFUSEDXML, LXML
from fastpyxl.workbook import Workbook
from fastpyxl.reader.excel import load_workbook as open
from fastpyxl.reader.excel import load_workbook
import fastpyxl._constants as constants

# Expose constants especially the version number

__author__ = constants.__author__
__author_email__ = constants.__author_email__
__license__ = constants.__license__
__maintainer_email__ = constants.__maintainer_email__
__url__ = constants.__url__
__version__ = constants.__version__
