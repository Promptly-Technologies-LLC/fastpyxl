# Copyright (c) 2010-2024 fastpyxl

"""
math.prod equivalent for < Python 3.8
"""

import functools
import operator

from typing import Any


def product(sequence):
    return functools.reduce(operator.mul, sequence)


prod: Any
try:
    from math import prod
except ImportError:
    prod = product
