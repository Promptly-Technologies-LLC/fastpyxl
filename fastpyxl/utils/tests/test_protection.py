# Copyright (c) 2010-2024 fastpyxl

from .. protection import hash_password


def test_password():
    enc = hash_password('secret')
    assert enc == 'DAA7'
