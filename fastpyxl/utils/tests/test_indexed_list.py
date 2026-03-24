import pytest


@pytest.fixture
def list():
    from ..indexed_list import IndexedList
    return IndexedList


def test_ctor(list):
    idx_list = list(['b', 'a'])
    assert idx_list == ['b', 'a']
    assert idx_list.clean is False


def test_allow_duplicate_ctor(list):
    idx_list = list(['b', 'a', 'b'])
    assert idx_list == ['b', 'a', 'b']
    idx_list.append('a')
    assert idx_list == ['b', 'a', 'b']


def test_function(list):
    idx_list = list()
    idx_list.append('b')
    idx_list.append('a')
    assert idx_list == ['b', 'a']


def test_contains(list):
    idx_list = list(['a', 'b', 'a'])
    assert idx_list.clean is False
    assert 'a' in idx_list
    assert idx_list.clean is True


def test_index(list):
    idx_list = list(['a', 'b'])
    idx_list.append('a')
    assert idx_list == ['a', 'b']
    idx_list.append('c')
    assert idx_list.index('c') == 2
    assert idx_list.clean is True


def test_table_builder(list):
    sb = list()
    result = {'a':0, 'b':1, 'c':2, 'd':3}

    for letter in sorted(result):
        for x in range(5):
            sb.append(letter)
        assert sb.index(letter) == result[letter]
    assert sb == ['a', 'b', 'c', 'd']
