#!/usr/bin/env python
# -*- coding: utf-8 -*-
from os.path import dirname, abspath, join
import xlrd
from excellent import Writer
from excellent.backends import XL
from sure import expect


LOCAL_FILE = lambda *path: join(abspath(dirname(__file__)), *path)


def assert_first_sheets_are_the_same(file1, file2):
    wb1 = xlrd.open_workbook(file1)
    wb2 = xlrd.open_workbook(file2)

    sh1 = wb1.sheet_by_index(0)
    sh2 = wb2.sheet_by_index(0)

    get_cell = lambda cell: (cell.value, cell.ctype)
    get_cell_content = lambda row: map(get_cell, row)
    get_rows = lambda sheet: map(get_cell_content, map(sheet.row, range(sheet.nrows)))

    expect(get_rows(sh1)).to.equal(get_rows(sh2))


def test_write_data_with_headers_to_xl():
    "Writer should write data with header correctly in XL format"
    # Given a temporary file
    tmpfile = open('/tmp/test_1.xls', 'w+')

    # Given an XL Writer
    writer = Writer(backend=XL(), output=tmpfile)

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025}]
    writer.write(data)

    # And we save it
    writer.save()

    # Then the written to data should match our expectation
    assert_first_sheets_are_the_same('/tmp/test_1.xls', LOCAL_FILE('excel1.xls'))
