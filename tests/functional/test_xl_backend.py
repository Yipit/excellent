#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Copyright <2013> Gabriel Falcao <gabriel@yipit.com>
# Copyright <2013> Suneel Chakravorty <suneel@yipit.com>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
import os
from os.path import dirname, abspath, join
from xlwt import XFStyle
import xlrd
from collections import OrderedDict
from excellent import Writer, XL
from excellent.exceptions import TooManyRowsError
from sure import expect, scenario


LOCAL_FILE = lambda *path: join(abspath(dirname(__file__)), *path)


def with_tmp_file(context):
    # Given a temporary file
    file_path = '/tmp/test.xls'
    try:
        # Delete file if it already exists
        os.remove(file_path)
    except OSError:
        pass
    context.tmpfile = open(file_path, 'w+')


def assert_first_sheets_are_the_same(file1, file2):
    wb1 = xlrd.open_workbook(file1, formatting_info=True)
    wb2 = xlrd.open_workbook(file2, formatting_info=True)

    assert wb1._all_sheets_count == wb2._all_sheets_count, (
        "\nThe number of sheets in {} is {} \n"
        "while the number of sheets in {} is {}".
        format(file1, wb1._all_sheets_count, file2, wb2._all_sheets_count))

    for index in range(wb1._all_sheets_count):
        sheet1 = wb1.sheet_by_index(index)
        sheet2 = wb2.sheet_by_index(index)

        assert sheet1.name == sheet2.name, (
            "\nSheet index {index} in file {f1} has name `{n1}`\n"
            " while \n"
            "sheet index {index} in file {f2} has name `{n2}`".
            format(index=index, f1=file1, n1=sheet1.name, f2=file2, n2=sheet2.name))

        def get_cell_content(sheet, cell):
            workbook = sheet.book
            cell_style = workbook.xf_list[cell.xf_index]
            format_type = workbook.format_map[cell_style.format_key]
            font_type = workbook.font_list[cell_style.font_index]

            # Normalize these since they are redundant and often set wrong
            # https://github.com/python-excel/xlrd/blob/master/xlrd/formatting.py#L184
            font_type.bold = 0
            font_type.family = 0
            font_type.character_set = 0

            # Normalize this since it will be different for each workbook
            font_type.font_index = 0

            # Normalize format_type attrs
            format_type.format_key = 0
            format_type.type = 0
            format_type.format_str = format_type.format_str.lower()

            return (
                cell.ctype,
                cell.value,
                cell_style.alignment,
                cell_style.background,
                cell_style.border,
                font_type,
                format_type,
                cell_style.is_style
            )

        def get_row_content(sheet, row):
            return [get_cell_content(sheet, cell) for cell in row]

        def get_rows(sheet):
            rows = [sheet.row(index) for index in range(sheet.nrows)]
            return [get_row_content(sheet, row) for row in rows]

        try:
            expect(get_rows(sheet1)).to.equal(get_rows(sheet2))
        except AssertionError as e:
            raise AssertionError(
                "In the sheet name {0}\n{1}".format(sheet1.name, e.message))


@scenario(with_tmp_file)
def test_write_data_with_headers_to_xl(context):
    "Writer should write data with header correctly in XL format"

    # Given an XL Writer
    writer = Writer(backend=XL(), output=context.tmpfile)

    # When we write some data
    data = [
        [("Country", "Argentina"), ("Revenue", 14500025)]
    ]
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
    data = [
        [("Country", "Argentina"), ("Revenue", 14500025)]
    ]
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
    data = [
        [("Country", "Argentina"), ("Revenue", 14500025)]
    ]
    writer.write(data)

    # When we write data to Awesome Sheet2
    backend.use_sheet('Awesome Sheet2')
    data = [
        OrderedDict([("Country", "Puerto Rico"), ("Revenue", 2340982)]),
        OrderedDict([("Country", "Colombia"), ("Revenue", 23409822)]),
        OrderedDict([("Country", "Brazil"), ("Revenue", 19982793)]),
    ]
    writer.write(data)

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet2.xls'))


@scenario(with_tmp_file)
def test_writing_multiple_times_to_same_sheet_and_multiple_sheets(context):
    "Writer can switch between sheets and write multiple times to same sheet"
    # Given a backend
    backend = XL()

    # And a writer
    writer = Writer(backend=backend, output=context.tmpfile)

    # And then switch to Awesome Sheet1 and this just adds this sheet
    backend.use_sheet('Awesome Sheet1')

    # When we write data to Awesome Sheet2
    backend.use_sheet('Awesome Sheet2')
    data = [
        OrderedDict([("Country", "Puerto Rico"), ("Revenue", 2340982)]),
    ]
    writer.write(data)

    backend.use_sheet('Awesome Sheet1')
    data = [
        OrderedDict([("Country", "Argentina"), ("Revenue", 14500025)])
    ]
    writer.write(data)

    # And switch back to Awesome Sheet 2 to write more data
    backend.use_sheet('Awesome Sheet2')
    data = [
        OrderedDict([("Country", "Colombia"), ("Revenue", 23409822)]),
        OrderedDict([("Country", "Brazil"), ("Revenue", 19982793)]),
    ]
    writer.write(data)

    writer.save()

    # Then the written data to Awesome Sheets 1 and 2 should match
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet2.xls'))


@scenario(with_tmp_file)
def test_writing_to_same_sheet_multiple_times_without_no_sheet_name(context):
    "Writer able to write  any number of times to the same sheet with no sheet name"
    # Given a backend
    backend = XL()

    # And a writer
    writer = Writer(backend=backend, output=context.tmpfile)

    data = [
        [("Country", "Argentina"), ("Revenue", 14500025)]
    ]
    writer.write(data)

    data = [
        OrderedDict([("Country", "Puerto Rico"), ("Revenue", 2340982)]),
        OrderedDict([("Country", "Colombia"), ("Revenue", 23409822)]),
        OrderedDict([("Country", "Brazil"), ("Revenue", 19982793)]),
    ]
    writer.write(data)

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet3.xls'))


@scenario(with_tmp_file)
def test_writing_with_custom_style_works(context):
    "Writer should be able to take a custom style and use that style in writing"

    style = XFStyle()
    style.font.bold = True

    backend= XL()
    writer = Writer(backend=backend, output=context.tmpfile)

    writer.write(
        [[("Country", "Argentina"), ("Revenue", 14500025)]],
        style=style,
    )

    writer.write(
        [[("Country", "Puerto Rico"), ("Revenue", 2340982)]]
    )

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet4.xls'))


@scenario(with_tmp_file)
def test_writing_with_default_style_works(context):
    "Writer should be able to take a default style"

    style = XFStyle()
    style.font.bold = True

    backend= XL(default_style=style)
    writer = Writer(backend=backend, output=context.tmpfile)

    writer.write(
        [[("Country", "Argentina"), ("Revenue", 14500025)]],
    )

    writer.write(
        [[("Country", "Puerto Rico"), ("Revenue", 2340982)]],
        bold=False,
    )

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet4.xls'))


@scenario(with_tmp_file)
def test_writing_with_bold_works(context):
    "Writer should be able write things in bold"

    backend= XL()
    writer = Writer(backend=backend, output=context.tmpfile)

    writer.write(
        [[("Country", "Argentina"), ("Revenue", 14500025)]],
        bold=True,
    )

    writer.write(
        [[("Country", "Puerto Rico"), ("Revenue", 2340982)]]
    )

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet4.xls'))


@scenario(with_tmp_file)
def test_writing_with_bottom_border_works(context):
    "Writer should be able write things with bottom borders"

    backend= XL()
    writer = Writer(backend=backend, output=context.tmpfile)

    writer.write(
        [[("Country", "Argentina"), ("Revenue", 14500025)]],
    )

    writer.write(
        [[("Country", "Puerto Rico"), ("Revenue", 2340982)]],
        bottom_border=True,
    )

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet5.xls'))


@scenario(with_tmp_file)
def test_writing_with_format_strings(context):
    "Writer should be able write things with format strings"

    backend= XL()
    writer = Writer(backend=backend, output=context.tmpfile)

    writer.write(
        [[("Country", "Argentina"), ("Revenue", 14500025)]],
    )

    writer.write(
        [[("Country", "Puerto Rico"), ("Revenue", 2340982)]],
        format_string='#,##0',
    )

    writer.save()

    # Then the written data should be under Awesome Sheet
    assert_first_sheets_are_the_same(context.tmpfile.name, LOCAL_FILE('awesome_sheet6.xls'))


@scenario(with_tmp_file)
def test_writing_too_many_rows_raises_error(context):
    "Writer should raise an error if we try to write too many rows"

    backend= XL()
    writer = Writer(backend=backend, output=context.tmpfile)

    for index in range(65535):
        writer.write(
            [[("Country", "Argentina"), ("Revenue", 14500025)]],
        )

    (writer.write.when
        .called_with([[("Country", "Argentina"), ("Revenue", 14500025)]])
        .should.throw(TooManyRowsError))
