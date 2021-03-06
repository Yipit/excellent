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
from __future__ import unicode_literals

from collections import OrderedDict
from mock import Mock, call
from sure import expect
from excellent import XL
from excellent.backends.xl_backend import default_style, STYLE_CACHE, get_column_width
from xlwt import Alignment


def test_get_header_style():
    ("XL backend .get_header_style() returns "
     "bold and horizontal right alignment")

    xl_backend = XL()
    style = xl_backend.get_header_style()

    expect(style).to.be.a('xlwt.XFStyle')
    expect(style.alignment).to.be.an(Alignment)

    expect(style.alignment.horz).to.equal(Alignment.HORZ_RIGHT)

    expect(style.borders).to.be.a('xlwt.Borders')
    expect(style.font).to.be.a('xlwt.Font')
    expect(style.font.bold).to.be.true


def test_use_sheet():
    ("XL backend .use_sheet() sets both "
     "`current_sheet` and `current_row`")

    # Given a XL backend
    class MyXLBackend(XL):
        get_or_create_sheet = Mock(return_value=('some sheet', 42))

    xl_backend = MyXLBackend()

    # It starts with no current_sheet and current_row
    xl_backend.current_sheet.should.be.none
    xl_backend.current_row.should.equal(0)

    # When I call .use_sheet
    xl_backend.use_sheet('an awesome sheet')

    # Then get_or_create_sheet gets called with the given sheet name
    MyXLBackend.get_or_create_sheet.assert_called_once_with('an awesome sheet')

    # And the current_sheet and current_row gets set
    xl_backend.current_sheet.should.equal('some sheet')
    xl_backend.current_row.should.equal(42)


def test_get_or_create_sheet_creates_new_sheet():
    ("XL backend should create a new sheet when a sheet of that name"
     " does not exist yet")

    # Given a mocked workbook and XL backend with no sheets
    workbook = Mock()

    class MyEmptyXL(XL):
        def get_sheets(self):
            return []

    xl_backend = MyEmptyXL(workbook=workbook)

    # When I call get_or_create_sheet of `anysheet`
    current_sheet, current_row = xl_backend.get_or_create_sheet('anysheet')

    # Then workbook.add_sheet should have been called once with `anysheet`
    workbook.add_sheet.assert_called_once_with('anysheet')

    # And the current sheet is the result of workbook.add_sheet
    current_sheet.should.equal(workbook.add_sheet.return_value)

    # And current_row should be zero
    current_row.should.equal(0)


def test_get_or_create_sheet_gets_existing_sheet():
    ("XL backend should create a new sheet when a sheet of that name"
     " does not exist yet")

    # Given a mocked workbook and XL backend with one sheet
    workbook = Mock()
    sheet = Mock()
    sheet.name = 'awesomesheet'
    sheet.rows = {0: 'stuff', 1: 'more stuff'}

    class MyXL(XL):
        def get_sheets(self):
            return [sheet]

    xl_backend = MyXL(workbook=workbook)

    # When I call get_or_create_sheet of `awesome sheet1
    current_sheet, current_row = xl_backend.get_or_create_sheet('awesomesheet')

    # Then workbook.add_sheet should not have been called
    workbook.add_sheet.called.should.be.false

    # And the current sheet is equal to the mocked sheet
    current_sheet.should.equal(sheet)

    # And current_row should be 1
    current_row.should.equal(1)


def test_get_sheets():
    ("XL backend.get_sheets should return"
     " number of worksheets in the workbook")
    # Given a mocked workbook and XL backend
    workbook = Mock()
    xl_backend = XL(workbook=workbook)

    # When I get sheets
    sheets = xl_backend.get_sheets()

    # Then sheets should equal the sheets form the workbook
    worksheets = workbook._Workbook__worksheets
    sheets.should.equal(worksheets)


def test_save():
    "XL backend.save should save th workbook and close the output"

    # Given an XL backend with mocked workbook and a mocked output
    workbook = Mock()
    output = Mock()
    xl_backend = XL(workbook=workbook)

    # When I save the backend to output
    xl_backend.save(output)

    # Then xl_backend.workbook.save should have been called once with output
    workbook.save.assert_called_once_with(output)

    # And output.close should hae been called once
    output.close.assert_called_once_with()


def test_write_row_passing_style():
    ("XL backend.write_row should write row using appropriate "
     "index, value and style")
    # Given a row
    row = Mock()

    # When .write_row is called
    xl_backend = XL()
    xl_backend.write_row(row, ['Cell One', 'Cell Two'], default_style)

    # Then row.write is called appropriately
    row.write.assert_has_calls([
        call(0, 'Cell One', default_style),
        call(1, 'Cell Two', default_style),
    ])


def test_write_row_passing_no_style():
    ("XL backend.write_row should write row using appropriate "
     "index, value and style")
    row = Mock()

    xl_backend = XL()
    xl_backend.write_row(row, ['Cell One', 'Cell Two'])

    row.write.assert_has_calls([
        call(0, 'Cell One', default_style),
        call(1, 'Cell Two', default_style),
    ])


def test_write_with_no_current_sheet_creates_sheet():
    ("XL backend.write with no current_row should create "
     "a sheet named 'Sheet1'")

    class MyXLBackend(XL):
        get_sheets = Mock(return_value=[])

    # Given a workbook
    workbook = Mock()
    workbook.add_sheet.return_value.name = 'I WAS JUST CREATED'
    # When calling write right away
    backend = MyXLBackend(workbook)
    backend.write([], Mock())

    # Then the current sheet is a new one
    workbook.add_sheet.assert_called_once_with('Sheet1')
    backend.current_sheet.name.should.equal('I WAS JUST CREATED')


def test_write_with_no_current_sheet_and_no_current_row():
    ("XL backend.write with no current_row and no current_sheet "
     "should write the headers of the columns and its contents")

    class MyXLBackend(XL):
        write_row = Mock()
        get_header_style = Mock(return_value='a cool style')
        get_sheets = Mock(return_value=[])

    data = [
        {'Name': 'Chuck Norris', 'Power': 'unlimited'},
        {'Name': 'Steven Seagal', 'Power': 'break necks'},
    ]
    workbook = Mock()
    sheet = workbook.add_sheet.return_value
    sheet.row.side_effect = lambda index: "this is the row {}".format(index)
    output = Mock()

    backend = MyXLBackend(workbook)
    backend.write(data, output)

    MyXLBackend.write_row.assert_has_calls([
        call("this is the row 0", ["Name", "Power"], 'a cool style', header_row=True),
        call("this is the row 1", ["Chuck Norris", "unlimited"], style=None, header_row=False),
        call("this is the row 2", ["Steven Seagal", "break necks"], style=None, header_row=False),
    ])


def test_write_with_current_sheet_no_current_row_and_no_rows():
    ("XL backend.write with a `current_sheet` assigned "
     "should write the headers of the columns and its contents")

    # Given a sheet with no rows
    mocked_row = Mock(side_effect=lambda index: "mocked row {}".format(index))

    current_sheet = Mock()
    current_sheet.name = 'Useful Sheet'
    current_sheet.row = mocked_row
    current_sheet.rows = {}

    # And a backend containing that sheet
    class MyXLBackend(XL):
        write_row = Mock()
        get_header_style = Mock(return_value='a cool style')

        def get_sheets(self):
            return [current_sheet]

    # And some data
    data = [
        [('Name', 'Chuck Norris'), ('Power', 'unlimited')],
        [('Name', 'Steven Seagal'), ('Power', 'break necks')],
    ]
    data = map(OrderedDict, data)
    workbook = Mock()
    output = Mock()

    # When calling .write() after calling .use_sheet()
    backend = MyXLBackend(workbook)
    backend.use_sheet('Useful Sheet')
    backend.write(data, output)

    # Then data should be written to the existing row
    MyXLBackend.write_row.assert_has_calls([
        call("mocked row 0", ["Name", "Power"], 'a cool style', header_row=True),
        call("mocked row 1", ["Chuck Norris", "unlimited"], style=None, header_row=False),
        call("mocked row 2", ["Steven Seagal", "break necks"], style=None, header_row=False),
    ])


def test_write_with_current_sheet_and_current_row():
    ("XL backend.write with a `current_sheet` and `current_row` assigned "
     "should write the headers of the columns and its contents")

    # Given a sheet with no rows
    mocked_row = Mock(side_effect=lambda index: "mocked row {}".format(index))

    current_sheet = Mock()
    current_sheet.name = 'Useful Sheet'
    current_sheet.row = mocked_row
    current_sheet.rows = {x: x for x in range(10)}

    # And a backend containing that sheet
    class MyXLBackend(XL):
        write_row = Mock()
        get_header_style = Mock(return_value='a cool style')

        def get_sheets(self):
            return [current_sheet]

    # And some data
    data = [
        OrderedDict([('Name', 'Chuck Norris'), ('Power', 'unlimited')]),
        OrderedDict([('Name', 'Steven Seagal'), ('Power', 'break necks')]),
    ]
    workbook = Mock()
    output = Mock()

    # When calling .write() after calling .use_sheet()
    backend = MyXLBackend(workbook)
    backend.current_row = 9
    backend.use_sheet('Useful Sheet')
    backend.write(data, output)

    # Then data should be written to the existing row
    MyXLBackend.write_row.assert_has_calls([
        call("mocked row 10", ["Chuck Norris", "unlimited"], style=None, header_row=False),
        call("mocked row 11", ["Steven Seagal", "break necks"], style=None, header_row=False),
    ])


def test_xl_style_cache_works():
    ("The XL backend should cache styles that are the same instead of creating "
    "a new style each time")

    STYLE_CACHE.clear()

    backend = XL()
    output = Mock()

    mock_row = Mock()

    backend.write_row(
        mock_row,
        [['Chuck Norris', 'Power']],
        output=output,
        bold=True,
    )

    STYLE_CACHE.should.have.length_of(1)
    bold_style = STYLE_CACHE.values()[0]

    backend.write_row(
        mock_row,
        [['Name', 'Some other guy']],
        output=output,
    )

    STYLE_CACHE.should.have.length_of(2)
    non_bold_style = [style for style in STYLE_CACHE.values() if style != bold_style][0]

    backend.write_row(
        mock_row,
        [['Name', 'Another person']],
        output=output,
        bold=True,
    )

    mock_row.write.assert_has_calls([
        call(0, ['Chuck Norris', 'Power'], bold_style),
        call(0, ['Name', 'Some other guy'], non_bold_style),
        call(0, ['Name', 'Another person'], bold_style),
    ])

    STYLE_CACHE.should.have.length_of(2)


def test_writing_with_none_format_string():
    "Writer should not use a format string if None is passed"

    STYLE_CACHE.clear()

    backend = XL()
    output = Mock()

    mock_row = Mock()

    backend.write_row(
        mock_row,
        [['Chuck Norris', 'Power']],
        format_string=None,
    )

    STYLE_CACHE.should.have.length_of(1)
    active_style = STYLE_CACHE.values()[0]
    active_style.num_format_str.should.equal("General")


def test_get_column_width():
    "Get column width should return the size of the 'value' for that column or the miniumum width, whichever is larger."
    get_column_width("a"*10).should.equal(3000)
    get_column_width("a"*100).should.equal(27500)
