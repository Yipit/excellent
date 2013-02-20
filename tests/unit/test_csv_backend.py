#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from mock import Mock
from sure import expect
from excellent import XL
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
    # workbook.add_sheet.assert_not_called_once_with('awesomesheet')

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
