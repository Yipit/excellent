#!/usr/bin/env python
# -*- coding: utf-8 -*-
from os.path import dirname, abspath, join
import xlrd
from excellent import Writer
from excellent.backends import XL
from sure import expect, scenario


LOCAL_FILE = lambda *path: join(abspath(dirname(__file__)), *path)


def with_tmp_file(context):
    # Given a temporary file
    context.tmpfile = open('/tmp/test.xls', 'w+')


def assert_first_sheets_are_the_same(file1, file2):
    wb1 = xlrd.open_workbook(file1)
    wb2 = xlrd.open_workbook(file2)

    assert wb1._all_sheets_count == wb2._all_sheets_count, (
        "\nThe number of sheets in {} is {} \n"
        "while the number of sheets in {} is {}".
        format(file1, wb1._all_sheets_count, file2, wb2._all_sheets_count))

    for index in range(wb1._all_sheets_count):
        sh1 = wb1.sheet_by_index(index)
        sh2 = wb2.sheet_by_index(index)

        assert sh1.name == sh2.name, (
            "\nSheet index {index} in file {f1} has name `{n1}`\n"
            " while \n"
            "sheet index {index} in file {f2} has name `{n2}`".
            format(index=index, f1=file1, n1=sh1.name, f2=file2, n2=sh2.name))

        get_cell = lambda cell: (cell.value, cell.ctype)
        get_cell_content = lambda row: map(get_cell, row)

        get_rows = lambda sheet: map(get_cell_content,
                                     map(sheet.row, range(sheet.nrows)))

        expect(get_rows(sh1)).to.equal(get_rows(sh2))


@scenario(with_tmp_file)
def test_write_data_with_headers_to_xl(context):
    "Writer should write data with header correctly in XL format"

    # Given an XL Writer
    writer = Writer(backend=XL(), output=context.tmpfile)

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025}]
    writer.write(data)

    # And we save it
    writer.save()

    # Then the written to data should match our expectation
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('excel1.xls'))


@scenario(with_tmp_file)
def test_write_data_to_xl_specifying_sheet_name(context):
    "Writer should write data to the sheet specified"
    # Given a backend
    backend = XL()

    # And a writer
    writer = Writer(backend=backend, output=context.tmpfile)

    # When we use sheet `Awesome Sheet1`
    backend.use_sheet('Awesome Sheet1')

    # And  we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025}]
    writer.write(data)
    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet1.xls'))


@scenario(with_tmp_file)
def test_writing_to_multiple_sheets_in_same_book(context):
    "Writer should write to multiple sheets in the same book"
    # Given a backend
    backend = XL()

    # And a writer
    writer = Writer(backend=backend, output=context.tmpfile)

    # When we write data to Awesome Sheet1
    backend.use_sheet('Awesome Sheet1')
    data = [{"Country": "Argentina", "Revenue": 14500025}]
    writer.write(data)

    # When we write data to Awesome Sheet2
    backend.use_sheet('Awesome Sheet2')
    data = [{"Country": "Puerto Rico", "Revenue": 2340982},
            {"Country": "Colombia", "Revenue": 23409822},
            {"Country": "Brazil", "Revenue": 19982793}]
    writer.write(data)

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet2.xls'))
